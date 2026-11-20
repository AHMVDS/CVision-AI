import html
import json
import os
import re
from io import BytesIO

import streamlit as st
import streamlit.components.v1 as components
from docx import Document
from google import genai
from google.genai import types
from pypdf import PdfReader

st.set_page_config(
    page_title="CVision AI — Smart Resume Intelligence",
    page_icon="✦",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# =========================
# SESSION STATE
# =========================

if "lang" not in st.session_state:
    st.session_state.lang = "ar"

if "result" not in st.session_state:
    st.session_state.result = None

if "result_lang" not in st.session_state:
    st.session_state.result_lang = None

if "cv_text" not in st.session_state:
    st.session_state.cv_text = ""

if "file_name" not in st.session_state:
    st.session_state.file_name = ""


# =========================
# TRANSLATIONS
# =========================

TEXT = {
    "ar": {
        "switch_label": "EN",
        "switch_help": "Switch interface to English",
        "subtitle": "منصة ذكية لتحليل وتقييم السير الذاتية بمعايير التوظيف الحديثة",
        "badge": "ذكاء اصطناعي • تقييم ATS واحترافي",
        "hero": "طوّر سيرتك الذاتية",
        "hero_highlight": "لتنافس بثقة.",
        "description": "احصل على تقييم فوري ودقيق لمحتوى سيرتك الذاتية، واكتشف نقاط القوة والمهارات البارزة والفجوات التي تحتاج إلى تحسين قبل التقديم على الوظائف.",
        "upload": "ارفع سيرتك الذاتية لبدء التحليل",
        "upload_hint": "ندعم ملفات PDF وDOCX حتى حجم 10 ميجابايت مع استخراج سريع للنص وتحليل منظم.",
        "drop": "اسحب ملف السيرة الذاتية هنا أو اضغط لاختيار ملف من جهازك",
        "formats": "الصيغ المدعومة: PDF, DOCX • الحد الأقصى: 10 MB",
        "analyze": "ابدأ التحليل الذكي الآن",
        "loading": "جاري تحليل السيرة الذاتية وتقييم جاهزيتها...",
        "retranslating": "جاري تحديث نتائج التحليل إلى اللغة العربية...",
        "score": "مؤشر جاهزية السيرة الذاتية",
        "summary": "الملخص التنفيذي للتقييم",
        "skills": "المهارات المكتشفة",
        "missing": "عناصر تحتاج إلى تحسين",
        "strengths": "أبرز نقاط القوة",
        "suggestions": "توصيات عملية للتطوير",
        "new": "تحليل سيرة ذاتية أخرى",
        "error_size": "حجم الملف يتجاوز الحد المسموح (10 ميجابايت). يرجى اختيار ملف أصغر.",
        "error_text": "تعذر استخراج نص واضح من الملف. تأكد أن الملف يحتوي على نص قابل للقراءة وليس صورًا ممسوحة فقط.",
        "error_api": "تعذر إتمام التحليل حاليًا. تحقق من إعداد مفتاح GEMINI_API_KEY وحاول مجددًا.",
        "footer": "خطوة أذكى وأكثر احترافية نحو فرصتك المهنية القادمة.",
        "f1": "1. رفع الملف بسهولة",
        "f1d": "ارفع ملف PDF أو DOCX خلال ثوانٍ دون خطوات معقدة.",
        "f2": "2. تحليل ذكي شامل",
        "f2d": "قراءة دقيقة للخبرات والمهارات وملاءمة هيكلة السيرة الذاتية.",
        "f3": "3. توصيات قابلة للتنفيذ",
        "f3d": "خطوات واضحة لتحسين الصياغة وسد الفجوات ورفع فرص القبول.",
        "step": "كيف تعمل المنصة في 3 خطوات",
        "result_label": "تقرير التحليل الذكي",
        "result_title": "نتائج تقييم السيرة الذاتية",
        "excellent": "جاهزية ممتازة",
        "good": "مستوى جيد جدًا",
        "average": "يحتاج إلى تطوير",
        "weak": "يحتاج إلى تحسين مكثف",
        "empty": "لا توجد عناصر مسجلة في هذا القسم.",
        "file_ready": "الملف جاهز للتحليل",
        "file_size_label": "الحجم",
        "secure": "معالجة سريعة • معايير ATS • عرض واضح ومنظم",
        "analysis_note": "ملاحظة: هذا التقييم استرشادي لقياس جودة السيرة الذاتية ووضوحها، ولا يُعد ضمانًا للقبول الوظيفي.",
        "items_count": "عناصر",
    },
    "en": {
        "switch_label": "عربي",
        "switch_help": "تبديل الواجهة إلى اللغة العربية",
        "subtitle": "Intelligent resume evaluation & ATS readiness platform",
        "badge": "ARTIFICIAL INTELLIGENCE • ATS & CAREER INSIGHTS",
        "hero": "Make your resume",
        "hero_highlight": "stand out.",
        "description": "Get an instant, structured evaluation of your CV to uncover core strengths, highlight key skills, and fix missing details before applying.",
        "upload": "Upload Your Resume for Analysis",
        "upload_hint": "Supports PDF and DOCX files up to 10 MB with fast text extraction and structured insights.",
        "drop": "Drag and drop your CV here or click to browse files",
        "formats": "Supported formats: PDF, DOCX • Max size: 10 MB",
        "analyze": "Run Smart CV Analysis",
        "loading": "Analyzing your resume and evaluating readiness...",
        "retranslating": "Updating analysis results into English...",
        "score": "CV Readiness Score",
        "summary": "Executive Evaluation Summary",
        "skills": "Identified Skills",
        "missing": "Areas to Improve",
        "strengths": "Key Strengths",
        "suggestions": "Actionable Recommendations",
        "new": "Analyze Another Resume",
        "error_size": "File size exceeds the 10 MB limit. Please upload a smaller file.",
        "error_text": "Could not extract readable text from this file. Make sure it contains selectable text rather than scanned images only.",
        "error_api": "Unable to complete analysis right now. Please check your GEMINI_API_KEY configuration and try again.",
        "footer": "A smarter, sharper step toward your next career milestone.",
        "f1": "1. Effortless Upload",
        "f1d": "Upload your PDF or DOCX resume in seconds with zero friction.",
        "f2": "2. Smart AI Evaluation",
        "f2d": "Objective review of skills, clarity, structure, and ATS readiness.",
        "f3": "3. Actionable Fixes",
        "f3d": "Practical improvements to strengthen impact and close gaps.",
        "step": "HOW IT WORKS IN 3 SIMPLE STEPS",
        "result_label": "INTELLIGENT CV REPORT",
        "result_title": "Resume Evaluation Results",
        "excellent": "Excellent Readiness",
        "good": "Very Good",
        "average": "Needs Development",
        "weak": "Needs Improvement",
        "empty": "No items to display in this section.",
        "file_ready": "File ready for analysis",
        "file_size_label": "Size",
        "secure": "Fast processing • ATS-minded review • Clean structured report",
        "analysis_note": "Disclaimer: This score provides indicative guidance on CV quality and clarity and does not guarantee job placement.",
        "items_count": "items",
    },
}


# =========================
# FILE EXTRACTION & AI ANALYSIS
# =========================

def extract_text(file_bytes: bytes, filename: str) -> str:
    lower_name = filename.lower()
    if lower_name.endswith(".pdf"):
        reader = PdfReader(BytesIO(file_bytes))
        return "\n".join((page.extract_text() or "") for page in reader.pages).strip()

    if lower_name.endswith(".docx"):
        doc = Document(BytesIO(file_bytes))
        parts = [p.text for p in doc.paragraphs if p.text and p.text.strip()]
        for table in doc.tables:
            for row in table.rows:
                row_vals = [cell.text.strip() for cell in row.cells if cell.text and cell.text.strip()]
                if row_vals:
                    parts.append(" | ".join(row_vals))
        return "\n".join(parts).strip()

    raise ValueError("Unsupported file format")


def analyze_cv(cv_text: str, target_lang_code: str) -> dict:
    api_key = st.secrets.get("GEMINI_API_KEY") or os.getenv("GEMINI_API_KEY")
    if not api_key:
        raise ValueError("Missing GEMINI_API_KEY")

    client = genai.Client(api_key=api_key)
    target_language = "Arabic" if target_lang_code == "ar" else "English"

    prompt = f"""
You are a senior HR recruitment specialist and ATS resume evaluator.
Analyze the following CV objectively, clearly, and concisely.

Write ALL text values strictly in {target_language}.
Return ONLY valid JSON matching this exact schema:
{{
  "compatibilityScore": 78,
  "summary": "Concise overall evaluation of structure, clarity, experience presentation, and ATS readiness.",
  "matchingSkills": ["Skill 1", "Skill 2"],
  "missingSkills": ["Missing or unclear CV element 1"],
  "strengths": ["Strength 1", "Strength 2"],
  "suggestions": ["Practical recommendation 1", "Practical recommendation 2"]
}}

Rules:
- Score must be an integer from 0 to 100.
- Do not fabricate personal facts, metrics, or experiences not present in the CV.
- Do not assume a narrow job title unless clearly indicated.
- Keep each list item crisp, specific, and actionable.
- Output JSON only.

CV CONTENT:
{cv_text[:18000]}
"""

    response = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=prompt,
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            temperature=0.25,
        ),
    )

    raw = (response.text or "").strip()
    raw = re.sub(r"^```(?:json)?\s*", "", raw)
    raw = re.sub(r"\s*```$", "", raw)
    return json.loads(raw)


def toggle_language():
    st.session_state.lang = "en" if st.session_state.lang == "ar" else "ar"


# =========================
# ACTIVE LANGUAGE & DIRECTION
# =========================

lang = st.session_state.lang
ar = lang == "ar"
t = TEXT[lang]
font = "Cairo" if ar else "Inter"
direction = "rtl" if ar else "ltr"
text_align = "right" if ar else "left"
opposite_align = "left" if ar else "right"

# Auto-refresh existing result into newly selected language when CV text is available
if (
    st.session_state.result is not None
    and st.session_state.result_lang != lang
    and st.session_state.cv_text.strip()
):
    try:
        with st.spinner(t["retranslating"]):
            st.session_state.result = analyze_cv(st.session_state.cv_text, lang)
            st.session_state.result_lang = lang
    except Exception:
        pass


# =========================
# DOCUMENT RTL / LTR SYNC
# =========================

components.html(
    f"""
    <script>
        (function() {{
            const doc = window.parent.document;
            if (doc && doc.documentElement) {{
                doc.documentElement.setAttribute("dir", "{direction}");
                doc.documentElement.setAttribute("lang", "{lang}");
            }}
            const appContainer = doc.querySelector('[data-testid="stAppViewContainer"]');
            if (appContainer) {{
                appContainer.setAttribute("dir", "{direction}");
            }}
        }})();
    </script>
    """,
    height=0,
)


# =========================
# PREMIUM BLUE & WHITE SAAS THEME
# =========================

st.markdown(
    f"""
    <style>
    @import url('
