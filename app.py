
import streamlit as st
import json
import os
import re
from io import BytesIO
from pypdf import PdfReader
from docx import Document
from google import genai
from google.genai import types

st.set_page_config(
    page_title="CVision AI",
    page_icon="📄",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# =========================
# SESSION STATE
# =========================

if "lang" not in st.session_state:
    st.session_state.lang = "ar"

if "result" not in st.session_state:
    st.session_state.result = None

if "cv_text" not in st.session_state:
    st.session_state.cv_text = ""

# =========================
# TRANSLATIONS
# =========================

TEXT = {
    "ar": {
        "subtitle": "نظام السيرة الذاتية الذكي للأعمال",
        "badge": "تقييم ذكي وفوري للسيرة الذاتية",
        "hero": "قيّم السيرة الذاتية بوضوح وسرعة",
        "description": "ارفع ملف السيرة الذاتية لتحليل محتواه وتقييم المهارات والخبرات وعرض نقاط القوة واقتراحات التحسين.",
        "start": "ابدأ الآن",
        "upload": "رفع السيرة الذاتية",
        "upload_hint": "الصيغ المدعومة PDF أو DOCX",
        "drop": "اختر ملف السيرة الذاتية",
        "size": "الحد الأقصى 10 ميجابايت",
        "analyze": "تحليل السيرة الذاتية",
        "loading": "جاري تحليل السيرة الذاتية...",
        "score": "نسبة التقييم والجاهزية",
        "summary": "التقييم العام",
        "skills": "المهارات المكتشفة",
        "missing": "المهارات والعناصر الناقصة",
        "strengths": "نقاط القوة",
        "suggestions": "اقتراحات للتحسين",
        "new": "تحليل سيرة ذاتية أخرى",
        "error_file": "يرجى اختيار ملف PDF أو DOCX.",
        "error_size": "حجم الملف أكبر من 10 ميجابايت.",
        "error_text": "تعذر استخراج النص من الملف.",
        "error_api": "تعذر إتمام التحليل. تأكد من إعدادات الخدمة وحاول مرة أخرى.",
        "footer": "CVision AI — نظام السيرة الذاتية الذكي",
        "f1": "رفع السيرة الذاتية",
        "f2": "تحليل ذكي",
        "f3": "نتائج واضحة وعملية",
        "f1d": "دعم ملفات PDF وDOCX واستخراج النص.",
        "f2d": "تحليل الخبرات والمهارات والمحتوى.",
        "f3d": "تقييم شامل وتوصيات للتحسين."
    },
    "en": {
        "subtitle": "Smart Business CV System",
        "badge": "Smart and Instant CV Evaluation",
        "hero": "Evaluate Resumes Clearly & Fast",
        "description": "Upload your CV to analyze its content, evaluate skills and experience, and receive strengths and improvement suggestions.",
        "start": "Get Started",
        "upload": "Upload Your CV",
        "upload_hint": "Supported formats: PDF or DOCX",
        "drop": "Choose your CV file",
        "size": "Maximum file size: 10 MB",
        "analyze": "Analyze CV",
        "loading": "Analyzing your CV...",
        "score": "CV Readiness Score",
        "summary": "Overall Evaluation",
        "skills": "Identified Skills",
        "missing": "Missing Skills & Elements",
        "strengths": "Candidate Strengths",
        "suggestions": "Improvement Suggestions",
        "new": "Analyze Another CV",
        "error_file": "Please upload a PDF or DOCX file.",
        "error_size": "File size exceeds 10 MB.",
        "error_text": "Could not extract readable text from the file.",
        "error_api": "Analysis failed. Please check the service configuration and try again.",
        "footer": "CVision AI — Smart Business CV System",
        "f1": "Upload CV",
        "f2": "Smart Analysis",
        "f3": "Clear Results",
        "f1d": "PDF and DOCX support with text extraction.",
        "f2d": "Analyze skills, experience and content.",
        "f3d": "Overall evaluation and recommendations."
    }
}

t = TEXT[st.session_state.lang]
ar = st.session_state.lang == "ar"

# =========================
# DESIGN
# =========================

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;500;600;700;800&family=Inter:wght@400;500;600;700;800&display=swap');

.stApp {{
    background: #f8fafc;
    color: #0f172a;
}}

html, body, [class*="css"] {{
    font-family: {'Cairo' if ar else 'Inter'}, sans-serif;
}}

[data-testid="stHeader"] {{
    background: transparent;
}}

.block-container {{
    max-width: 1000px;
    padding-top: 2rem;
    padding-bottom: 3rem;
}}

h1, h2, h3 {{
    color: #0f172a;
}}

.hero {{
    background: white;
    border: 1px solid #e2e8f0;
    border-radius: 25px;
    padding: 45px 30px;
    text-align: center;
    box-shadow: 0 10px 30px -10px rgba(37,99,235,.12);
    margin: 15px 0 25px;
}}

.logo {{
    color: #2563eb;
    font-size: 34px;
    font-weight: 800;
    letter-spacing: -1px;
}}

.badge {{
    display: inline-block;
    background: #eff6ff;
    color: #1d4ed8;
    border: 1px solid #dbeafe;
    padding: 7px 15px;
    border-radius: 30px;
    font-size: 13px;
    margin: 15px 0;
}}

.hero p {{
    color: #64748b;
    line-height: 1.9;
    font-size: 16px;
}}

.feature {{
    background: #f8fafc;
    border: 1px solid #e2e8f0;
    border-radius: 15px;
    padding: 17px;
    text-align: {'right' if ar else 'left'};
    margin-top: 15px;
}}

.feature strong {{
    color: #2563eb;
    display: block;
    margin-bottom: 5px;
}}

.feature span {{
    color: #64748b;
    font-size: 12px;
}}

.panel {{
    background: white;
    border: 1px solid #e2e8f0;
    border-radius: 20px;
    padding: 25px;
    margin-top: 20px;
}}

.score {{
    background: #eff6ff;
    border: 1px solid #bfdbfe;
    border-radius: 18px;
    padding: 20px;
    text-align: center;
}}

.score-number {{
    color: #2563eb;
    font-size: 45px;
    font-weight: 800;
}}

.result-box {{
    background: white;
    border: 1px solid #e2e8f0;
    border-radius: 16px;
    padding: 20px;
    margin: 12px 0;
}}

div.stButton > button {{
    background: #2563eb;
    color: white;
    border-radius: 11px;
    border: none;
    font-weight: 700;
    min-height: 45px;
}}

div.stButton > button:hover {{
    background: #1d4ed8;
    color: white;
}}

[data-testid="stFileUploader"] {{
    background: #f8fafc;
    border-radius: 15px;
}}

footer {{
    visibility: hidden;
}}

</style>
""", unsafe_allow_html=True)

# =========================
# LANGUAGE SWITCH
# =========================

col1, col2 = st.columns([3, 1])

with col2:
    language = st.selectbox(
        "Language",
        ["العربية", "English"],
        index=0 if ar else 1,
        label_visibility="collapsed"
    )

    selected = "ar" if language == "العربية" else "en"

    if selected != st.session_state.lang:
        st.session_state.lang = selected
        st.rerun()

# =========================
# HEADER
# =========================

st.markdown("""
<div style="text-align:center;margin:15px 0;">
    <div class="logo">CVision AI</div>
</div>
""", unsafe_allow_html=True)

st.caption(t["subtitle"])

# =========================
# HOME
# =========================

st.markdown(f"""
<div class="hero">
    <div class="badge">CVision AI • {t['badge']}</div>
    <h1>{t['hero']}</h1>
    <p>{t['description']}</p>
</div>
""", unsafe_allow_html=True)

f1, f2, f3 = st.columns(3)

for col, title, desc in [
    (f1, t["f1"], t["f1d"]),
    (f2, t["f2"], t["f2d"]),
    (f3, t["f3"], t["f3d"])
]:
    with col:
        st.markdown(
            f'<div class="feature"><strong>{title}</strong><span>{desc}</span></div>',
            unsafe_allow_html=True
        )

# =========================
# FILE EXTRACTION
# =========================

def extract_text(file):
    data = file.getvalue()

    if file.name.lower().endswith(".pdf"):
        reader = PdfReader(BytesIO(data))
        return "\n".join(
            page.extract_text() or "" for page in reader.pages
        )

    if file.name.lower().endswith(".docx"):
        doc = Document(BytesIO(data))
        return "\n".join(p.text for p in doc.paragraphs)

    raise ValueError("Unsupported file")


# =========================
# AI ANALYSIS
# =========================

def analyze_cv(cv_text):

    try:
        api_key = (
            st.secrets.get("GEMINI_API_KEY")
            or os.getenv("GEMINI_API_KEY")
        )

        if not api_key:
            raise ValueError("Missing API configuration")

        client = genai.Client(api_key=api_key)

        language = "Arabic" if ar else "English"

        prompt = f"""
You are a professional HR recruitment specialist and ATS CV evaluator.

Analyze this CV objectively and professionally.

Respond entirely in {language}.

Return ONLY valid JSON in this exact structure:

{{
 "compatibilityScore": 75,
 "summary": "Overall CV evaluation",
 "matchingSkills": ["Skill 1", "Skill 2"],
 "missingSkills": ["Missing item 1"],
 "strengths": ["Strength 1"],
 "suggestions": ["Suggestion 1"]
}}

Requirements:
- Score from 0 to 100.
- Evaluate CV quality, structure, experience, skills and ATS readiness.
- Do not invent candidate information.
- Do not assume a specific job unless one is provided.
- Explain missing or unclear CV elements.
- Give practical improvement recommendations.
- Keep results concise.
- Return JSON only.

CV CONTENT:
{cv_text[:18000]}
"""

        response = client.models.generate_content(
            model="gemini-2.5-flash",
            contents=prompt,
            config=types.GenerateContentConfig(
                response_mime_type="application/json",
                temperature=0.3
            )
        )

        raw = response.text.strip()
        raw = re.sub(r"^```(?:json)?\s*", "", raw)
        raw = re.sub(r"\s*```$", "", raw)

        return json.loads(raw)

    except Exception:
        raise RuntimeError("Analysis service unavailable")


# =========================
# UPLOAD SECTION
# =========================

st.markdown('<div class="panel">', unsafe_allow_html=True)

st.subheader("📄 " + t["upload"])
st.caption(t["upload_hint"])

uploaded_file = st.file_uploader(
    t["drop"],
    type=["pdf", "docx"],
    key="cv_upload"
)

if uploaded_file:

    if uploaded_file.size > 10 * 1024 * 1024:
        st.error(t["error_size"])
        st.stop()

    st.success(uploaded_file.name)

    if st.button(t["analyze"], use_container_width=True):

        try:
            with st.spinner(t["loading"]):

                cv_text = extract_text(uploaded_file)

                if not cv_text.strip() or len(cv_text.strip()) < 15:
                    st.error(t["error_text"])
                    st.stop()

                st.session_state.cv_text = cv_text
                st.session_state.result = analyze_cv(cv_text)

            st.rerun()

        except Exception:
            st.error(t["error_api"])

st.markdown('</div>', unsafe_allow_html=True)

# =========================
# RESULTS
# =========================

if st.session_state.result:

    result = st.session_state.result

    st.divider()
    st.subheader("📊 " + t["summary"])

    score = max(
        0,
        min(100, int(result.get("compatibilityScore", 0)))
    )

    st.markdown(f"""
    <div class="score">
        <div style="color:#1e40af;font-weight:700;">
            {t['score']}
        </div>
        <div class="score-number">{score}%</div>
    </div>
    """, unsafe_allow_html=True)

    st.progress(score / 100)

    st.markdown(
        f'<div class="result-box">{result.get("summary", "")}</div>',
        unsafe_allow_html=True
    )

    sections = [
        ("💡", "skills", "matchingSkills"),
        ("⚠️", "missing", "missingSkills"),
        ("⭐", "strengths", "strengths"),
        ("🚀", "suggestions", "suggestions")
    ]

    for icon, label, key in sections:

        st.markdown(f"### {icon} {t[label]}")

        items = result.get(key, [])

        if isinstance(items, list) and items:
            for item in items:
                st.markdown(f"- {item}")
        else:
            st.caption("—")

    st.divider()

    if st.button(t["new"], use_container_width=True):
        st.session_state.result = None
        st.session_state.cv_text = ""
        st.rerun()

# =========================
# FOOTER
# =========================

st.markdown("---")

st.markdown(
    f'<div style="text-align:center;color:#64748b;font-size:12px;">{t["footer"]}</div>',
    unsafe_allow_html=True
)

