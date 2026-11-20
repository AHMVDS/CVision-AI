import streamlit as st
import json
import os
from pypdf import PdfReader
from docx import Document
from google import genai
from google.genai import types

st.set_page_config(page_title="CVision AI", page_icon="📄", layout="centered")

if "language" not in st.session_state:
    st.session_state.language = "ar"

lang_col1, lang_col2 = st.columns(2)
with lang_col1:
    if st.button("العربية", use_container_width=True):
        st.session_state.language = "ar"
with lang_col2:
    if st.button("English", use_container_width=True):
        st.session_state.language = "en"

AR = st.session_state.language == "ar"

st.markdown("""
<style>
.stApp { background-color: #f5f8ff; }
.main-title { text-align:center; color:#174ea6; font-size:38px; font-weight:800; }
.subtitle { text-align:center; color:#64748b; font-size:16px; margin-bottom:30px; }
div.stButton > button { background-color:#1769e0; color:white; border-radius:10px; border:none; padding:10px 20px; }
div.stButton > button:hover { background-color:#0d47a1; color:white; }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">CVision AI</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="subtitle">حلّل سيرتك الذاتية وطور فرصك المهنية باستخدام الذكاء الاصطناعي</div>' if AR
    else '<div class="subtitle">Analyze your CV and improve your career opportunities with AI</div>',
    unsafe_allow_html=True,
)
st.divider()


def extract_text(file):
    filename = file.name.lower()
    if filename.endswith(".pdf"):
        reader = PdfReader(file)
        return "\n".join(page.extract_text() or "" for page in reader.pages)
    if filename.endswith(".docx"):
        document = Document(file)
        return "\n".join(paragraph.text for paragraph in document.paragraphs)
    return ""


def analyze_cv(cv_text, language):
    api_key = st.secrets.get("GEMINI_API_KEY") or os.getenv("GEMINI_API_KEY")
    if not api_key:
        raise ValueError("GEMINI_API_KEY is missing. Add it to Streamlit Secrets.")

    client = genai.Client(api_key=api_key)
    language_instruction = "Respond entirely in Arabic." if language == "ar" else "Respond entirely in English."
    prompt = f"""
You are a professional CV analyst and ATS recruitment expert.
Analyze the CV below. {language_instruction}
Return ONLY valid JSON with exactly this structure:
{{
  "score": 75,
  "summary": "A short overall evaluation",
  "skills": ["Skill 1", "Skill 2"],
  "missing_skills": ["Missing item 1"],
  "strengths": ["Strength 1", "Strength 2"],
  "suggestions": ["Suggestion 1", "Suggestion 2"]
}}
Rules:
- Score must be between 0 and 100.
- Evaluate clarity, structure, experience, education, skills and ATS compatibility.
- Do not invent information that is not in the CV.
- Missing skills should refer to important CV elements that are absent or unclear.
- Give practical and specific improvement suggestions.
- Keep the response concise.
- Return JSON only.

CV CONTENT:
{cv_text[:25000]}
"""
    response = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=prompt,
        config=types.GenerateContentConfig(response_mime_type="application/json", temperature=0.3),
    )
    return json.loads(response.text)


st.subheader("📄 ارفع سيرتك الذاتية" if AR else "📄 Upload Your CV")
st.write(
    "ارفع ملف PDF أو Word لتحليل محتواه وتقييمه." if AR
    else "Upload a PDF or Word document to analyze and evaluate its content."
)
uploaded_file = st.file_uploader("اختر الملف" if AR else "Choose your file", type=["pdf", "docx"])

if uploaded_file:
    file_size = uploaded_file.size / (1024 * 1024)
    if file_size > 10:
        st.error("حجم الملف يجب ألا يتجاوز 10MB." if AR else "File size must not exceed 10MB.")
    else:
        st.success(f"تم رفع الملف: {uploaded_file.name}" if AR else f"File uploaded: {uploaded_file.name}")
        if st.button("ابدأ تحليل السيرة الذاتية" if AR else "Analyze My CV", use_container_width=True):
            try:
                with st.spinner("جاري تحليل سيرتك الذاتية..." if AR else "Analyzing your CV..."):
                    cv_text = extract_text(uploaded_file)
                    if not cv_text.strip():
                        st.error(
                            "لم نتمكن من استخراج النص. تأكد أن الملف يحتوي على نص قابل للقراءة."
                            if AR else "No readable text found. Make sure your file contains selectable text."
                        )
                        st.stop()
                    result = analyze_cv(cv_text, "ar" if AR else "en")
                st.session_state.analysis_result = result
            except Exception as error:
                st.error(f"حدث خطأ أثناء التحليل: {error}" if AR else f"Analysis error: {error}")

if "analysis_result" in st.session_state:
    result = st.session_state.analysis_result
    st.divider()
    st.subheader("📊 نتيجة تحليل السيرة الذاتية" if AR else "📊 CV Analysis Results")
    score = max(0, min(100, int(result.get("score", 0))))
    st.metric("التقييم العام" if AR else "Overall Score", f"{score}/100")
    st.progress(score / 100)

    sections = [
        ("📝 الملخص" if AR else "📝 Summary", "summary"),
        ("💡 المهارات المكتشفة" if AR else "💡 Detected Skills", "skills"),
        ("⚠️ العناصر الناقصة" if AR else "⚠️ Missing Elements", "missing_skills"),
        ("⭐ نقاط القوة" if AR else "⭐ Strengths", "strengths"),
        ("🚀 اقتراحات التحسين" if AR else "🚀 Improvement Suggestions", "suggestions"),
    ]
    for heading, key in sections:
        st.subheader(heading)
        value = result.get(key, "")
        if isinstance(value, list):
            if value:
                for item in value:
                    st.write(f"• {item}")
            else:
                st.write("—")
        else:
            st.write(value or "—")

    st.divider()
    if st.button("تحليل ملف جديد" if AR else "Analyze Another CV", use_container_width=True):
        del st.session_state.analysis_result
        st.rerun()

st.markdown("---")
st.caption("CVision AI | AI-Powered CV Analysis")
