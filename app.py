import html
import json
import os
import re
from io import BytesIO
from pathlib import Path
from typing import List

import streamlit as st
import streamlit.components.v1 as components
from docx import Document
from google import genai
from google.genai import types
from pydantic import BaseModel, Field
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


def toggle_language():
    st.session_state.lang = "en" if st.session_state.lang == "ar" else "ar"


# =========================
# TRANSLATIONS
# =========================

TEXT = {
    "ar": {
        "switch_label": "🌐 EN",
        "switch_help": "Switch interface to English (LTR)",
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
        "error_key": "مفتاح الخدمة غير مضبوط في الإعدادات. يرجى إضافته في Secrets ثم إعادة تشغيل التطبيق.",
        "error_api": "تعذر إتمام التحليل حاليًا. يرجى التأكد من صلاحية المفتاح والاتصال ثم المحاولة مرة أخرى.",
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
        "switch_label": "🌐 عربي",
        "switch_help": "تبديل الواجهة إلى اللغة العربية (RTL)",
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
        "error_key": "Service key is not configured. Please add it to Secrets and reboot the app.",
        "error_api": "Unable to complete the analysis right now. Please verify your key and connection and try again.",
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
# BACKEND CONFIG, EXTRACTION & ENGINE
# =========================

KEY_CANDIDATES = ("CVISION_AI_KEY", "AI_SERVICE_KEY", "GEMINI_API_KEY", "GOOGLE_API_KEY")
MODEL_CANDIDATES = ("gemini-2.5-flash", "gemini-2.0-flash", "gemini-1.5-flash")


class CVEvaluationSchema(BaseModel):
    """Structured schema for CV evaluation output from Gemini."""

    compatibilityScore: int = Field(
        ...,
        ge=0,
        le=100,
        description="Overall CV quality, structure, clarity, and readiness score from 0 to 100.",
    )
    summary: str = Field(
        default="",
        description="Concise 2-3 sentence overall evaluation summary.",
    )
    matchingSkills: List[str] = Field(
        default_factory=list,
        description="Core skills and competencies identified in the CV.",
    )
    missingSkills: List[str] = Field(
        default_factory=list,
        description="Important skills, metrics, credentials, or sections missing or needing reinforcement.",
    )
    strengths: List[str] = Field(
        default_factory=list,
        description="Key candidate strengths evidenced in the CV.",
    )
    suggestions: List[str] = Field(
        default_factory=list,
        description="Actionable recommendations to improve the CV.",
    )


def _parse_secrets_toml_fallback() -> str:
    candidate_paths = [
        Path.cwd() / ".streamlit" / "secrets.toml",
        Path(__file__).resolve().parent / ".streamlit" / "secrets.toml",
    ]
    for path in candidate_paths:
        if not path.is_file():
            continue
        try:
            for raw_line in path.read_text(encoding="utf-8").splitlines():
                line = raw_line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                k, v = line.split("=", 1)
                if k.strip() in KEY_CANDIDATES:
                    cleaned = v.strip().strip('"').strip("'").strip()
                    if cleaned:
                        return cleaned
        except Exception:
            continue
    return ""


def get_service_key() -> str:
    for key_name in KEY_CANDIDATES:
        try:
            if key_name in st.secrets and st.secrets[key_name]:
                val = str(st.secrets[key_name]).strip()
                if val:
                    return val
        except Exception:
            pass

    for key_name in KEY_CANDIDATES:
        env_val = (os.getenv(key_name) or "").strip()
        if env_val:
            return env_val

    toml_val = _parse_secrets_toml_fallback()
    if toml_val:
        return toml_val

    return ""


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
    service_key = get_service_key()
    if not service_key:
        raise KeyError("MISSING_SERVICE_KEY")

    client = genai.Client(api_key=service_key)
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

    last_error = None
    for model_name in MODEL_CANDIDATES:
        try:
            response = client.models.generate_content(
                model=model_name,
                contents=prompt,
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    response_schema=CVEvaluationSchema,
                    temperature=0.25,
                ),
            )
            raw = (response.text or "").strip()
            raw = re.sub(r"^```(?:json)?\s*", "", raw, flags=re.IGNORECASE)
            raw = re.sub(r"\s*```$", "", raw, flags=re.IGNORECASE)
            parsed = json.loads(raw)
            if isinstance(parsed, dict):
                validated = CVEvaluationSchema(**parsed)
                return validated.model_dump()
        except Exception as exc:
            last_error = exc
            continue

    raise RuntimeError(str(last_error) if last_error else "ANALYSIS_FAILED")


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

if (
    st.session_state.result is not None
    and st.session_state.result_lang != lang
    and st.session_state.cv_text.strip()
    and get_service_key()
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
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;500;600;700;800;900&family=Inter:wght@400;500;600;700;800;900&display=swap');

    :root {{
        --blue-900: #0f295a;
        --blue-800: #163d82;
        --blue-700: #1d4ed8;
        --blue-600: #2563eb;
        --blue-500: #3b82f6;
        --blue-200: #bfdbfe;
        --blue-100: #dbeafe;
        --blue-50:  #eff6ff;
        --bg-top:   #f8fbff;
        --bg-bot:   #eef4fc;
        --surface:  #ffffff;
        --ink:      #0f2140;
        --muted:    #64748b;
        --border:   #e2ebf7;
        --shadow-sm: 0 6px 20px rgba(15, 41, 90, 0.05);
        --shadow-md: 0 14px 36px rgba(37, 99, 235, 0.09);
        --shadow-lg: 0 24px 60px rgba(15, 41, 90, 0.16);
    }}

    .stApp {{
        background:
            radial-gradient(circle at 8% 4%, rgba(59, 130, 246, 0.12), transparent 34%),
            radial-gradient(circle at 92% 14%, rgba(37, 99, 235, 0.10), transparent 32%),
            linear-gradient(180deg, var(--bg-top) 0%, var(--bg-bot) 100%);
        color: var(--ink);
        font-family: '{font}', sans-serif;
        direction: {direction};
    }}

    html, body, [class*="css"], .stMarkdown, .stText, label, input, textarea, button {{
        font-family: '{font}', sans-serif !important;
    }}

    [data-testid="stAppViewContainer"] {{
        direction: {direction};
    }}

    [data-testid="stHeader"] {{
        background: transparent;
        height: 2.5rem;
    }}

    .block-container {{
        max-width: 1160px;
        padding-top: 0.75rem;
        padding-bottom: 3rem;
    }}

    #MainMenu, footer, [data-testid="stToolbar"] {{
        visibility: hidden;
    }}

    /* TOP BAR */
    .topbar-card {{
        background: rgba(255, 255, 255, 0.90);
        backdrop-filter: blur(12px);
        border: 1px solid var(--border);
        border-radius: 20px;
        padding: 12px 20px;
        box-shadow: var(--shadow-sm);
        margin-bottom: 18px;
    }}

    .brand-inline {{
        display: flex;
        align-items: center;
        gap: 12px;
    }}

    .brand-badge {{
        width: 42px;
        height: 42px;
        border-radius: 13px;
        background: linear-gradient(135deg, var(--blue-600), var(--blue-500));
        color: #ffffff;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 19px;
        font-weight: 800;
        box-shadow: 0 8px 20px rgba(37, 99, 235, 0.25);
        flex-shrink: 0;
    }}

    .brand-title {{
        font-size: 21px;
        font-weight: 900;
        color: var(--blue-900);
        line-height: 1.15;
        letter-spacing: -0.4px;
    }}

    .brand-title span {{
        color: var(--blue-600);
    }}

    .brand-caption {{
        font-size: 12px;
        color: var(--muted);
        margin-top: 2px;
    }}

    /* REFINED LANGUAGE BUTTON STYLING */
    .lang-btn-box {{
        display: flex;
        justify-content: flex-end;
        align-items: center;
    }}

    .lang-btn-box div.stButton {{
        width: auto;
    }}

    .lang-btn-box div.stButton > button {{
        display: inline-flex !important;
        align-items: center !important;
        justify-content: center !important;
        gap: 6px !important;
        min-height: 38px !important;
        height: 38px !important;
        min-width: 84px !important;
        width: auto !important;
        padding: 0 16px !important;
        border-radius: 999px !important;
        background: linear-gradient(135deg, #ffffff 0%, var(--blue-50) 100%) !important;
        color: var(--blue-800) !important;
        border: 1.5px solid var(--blue-200) !important;
        font-size: 13px !important;
        font-weight: 800 !important;
        letter-spacing: 0.2px;
        box-shadow: 0 4px 12px rgba(37, 99, 235, 0.08) !important;
        transition: all 0.22s cubic-bezier(0.4, 0, 0.2, 1) !important;
        white-space: nowrap !important;
    }}

    .lang-btn-box div.stButton > button:hover {{
        background: linear-gradient(135deg, var(--blue-600) 0%, var(--blue-500) 100%) !important;
        color: #ffffff !important;
        border-color: var(--blue-600) !important;
        transform: translateY(-1.5px) scale(1.02) !important;
        box-shadow: 0 8px 20px rgba(37, 99, 235, 0.26) !important;
    }}

    .lang-btn-box div.stButton > button:active {{
        transform: translateY(0px) scale(0.98) !important;
        box-shadow: 0 3px 10px rgba(37, 99, 235, 0.2) !important;
    }}

    /* HERO BANNER */
    .hero-box {{
        position: relative;
        overflow: hidden;
        background:
            radial-gradient(circle at 85% 15%, rgba(147, 197, 253, 0.24), transparent 34%),
            radial-gradient(circle at 15% 85%, rgba(96, 165, 250, 0.18), transparent 35%),
            linear-gradient(125deg, var(--blue-900) 0%, var(--blue-800) 52%, var(--blue-600) 100%);
        border: 1px solid rgba(255, 255, 255, 0.16);
        border-radius: 28px;
        padding: 56px 36px;
        margin: 8px 0 28px;
        text-align: center;
        box-shadow: var(--shadow-lg);
    }}

    .hero-ring {{
        position: absolute;
        width: 340px;
        height: 340px;
        border-radius: 50%;
        border: 1px solid rgba(255, 255, 255, 0.14);
        top: -150px;
        {opposite_align}: -70px;
        pointer-events: none;
    }}

    .hero-inner {{
        position: relative;
        z-index: 2;
    }}

    .hero-pill {{
        display: inline-flex;
        align-items: center;
        gap: 6px;
        background: rgba(255, 255, 255, 0.14);
        border: 1px solid rgba(255, 255, 255, 0.24);
        color: #eff6ff;
        padding: 7px 16px;
        border-radius: 999px;
        font-size: 12px;
        font-weight: 800;
        margin-bottom: 18px;
        backdrop-filter: blur(8px);
    }}

    .hero-box h1 {{
        color: #ffffff;
        font-size: clamp(32px, 4.6vw, 52px);
        font-weight: 900;
        line-height: 1.28;
        margin: 0 0 14px;
    }}

    .hero-box h1 span {{
        color: #93c5fd;
    }}

    .hero-box p {{
        color: #dbeafe;
        font-size: 15px;
        line-height: 1.9;
        max-width: 720px;
        margin: 0 auto 24px;
    }}

    .hero-trust {{
        display: inline-block;
        background: rgba(15, 41, 90, 0.35);
        border: 1px solid rgba(255, 255, 255, 0.16);
        color: #e0f2fe;
        font-size: 12px;
        font-weight: 600;
        padding: 8px 18px;
        border-radius: 999px;
    }}

    /* SECTION KICKER */
    .kicker {{
        text-align: center;
        font-size: 12px;
        font-weight: 900;
        letter-spacing: 1.1px;
        color: var(--blue-700);
        margin: 26px 0 14px;
    }}

    /* FEATURE CARDS */
    .step-card {{
        background: var(--surface);
        border: 1px solid var(--border);
        border-radius: 22px;
        padding: 24px 20px;
        height: 100%;
        box-shadow: var(--shadow-sm);
        transition: transform 0.22s ease, box-shadow 0.22s ease, border-color 0.22s ease;
        text-align: {text_align};
    }}

    .step-card:hover {{
        transform: translateY(-4px);
        border-color: var(--blue-200);
        box-shadow: var(--shadow-md);
    }}

    .step-icon {{
        width: 46px;
        height: 46px;
        border-radius: 14px;
        background: linear-gradient(135deg, var(--blue-50), var(--blue-100));
        border: 1px solid var(--blue-200);
        color: var(--blue-700);
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 20px;
        margin-bottom: 14px;
    }}

    .step-card strong {{
        display: block;
        color: var(--blue-900);
        font-size: 16px;
        font-weight: 800;
        margin-bottom: 6px;
    }}

    .step-card span {{
        display: block;
        color: var(--muted);
        font-size: 13px;
        line-height: 1.75;
    }}

    /* UPLOAD SECTION */
    .upload-shell {{
        background: var(--surface);
        border: 1px solid var(--border);
        border-radius: 26px;
        padding: 28px 28px 18px;
        margin-top: 32px;
        box-shadow: var(--shadow-md);
    }}

    .upload-title {{
        font-size: 22px;
        font-weight: 900;
        color: var(--blue-900);
        margin-bottom: 6px;
        text-align: {text_align};
    }}

    .upload-sub {{
        font-size: 13.5px;
        color: var(--muted);
        margin-bottom: 18px;
        text-align: {text_align};
        line-height: 1.75;
    }}

    [data-testid="stFileUploader"] {{
        border: 2px dashed #93c5fd !important;
        border-radius: 20px !important;
        background: linear-gradient(180deg, #f8fbff 0%, #eff6ff 100%) !important;
        padding: 18px !important;
        transition: all 0.2s ease;
    }}

    [data-testid="stFileUploader"]:hover {{
        border-color: var(--blue-600) !important;
        background: #e8f1fe !important;
    }}

    [data-testid="stFileUploader"] section {{
        border: none !important;
        background: transparent !important;
    }}

    .file-meta-card {{
        background: var(--blue-50);
        border: 1px solid var(--blue-200);
        border-radius: 16px;
        padding: 14px 18px;
        margin: 14px 0 12px;
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 12px;
        flex-wrap: wrap;
    }}

    .file-meta-main {{
        display: flex;
        align-items: center;
        gap: 10px;
        color: var(--blue-900);
        font-weight: 800;
        font-size: 14px;
    }}

    .file-meta-tag {{
        background: #ffffff;
        color: var(--blue-700);
        border: 1px solid var(--blue-200);
        padding: 4px 12px;
        border-radius: 999px;
        font-size: 12px;
        font-weight: 700;
    }}

    /* PRIMARY ACTION BUTTONS */
    .action-btn-wrap div.stButton > button {{
        background: linear-gradient(115deg, var(--blue-700) 0%, var(--blue-600) 55%, var(--blue-500) 100%);
        color: #ffffff;
        border: none;
        border-radius: 15px;
        min-height: 50px;
        font-size: 15px;
        font-weight: 800;
        box-shadow: 0 10px 24px rgba(37, 99, 235, 0.24);
        transition: transform 0.2s ease, box-shadow 0.2s ease, filter 0.2s ease;
    }}

    .action-btn-wrap div.stButton > button:hover {{
        color: #ffffff;
        transform: translateY(-2px);
        filter: brightness(1.04);
        box-shadow: 0 14px 28px rgba(37, 99, 235, 0.32);
    }}

    [data-testid="stAlert"] {{
        border-radius: 15px;
        direction: {direction};
        text-align: {text_align};
    }}

    /* RESULTS SECTION */
    .report-banner {{
        background:
            radial-gradient(circle at 88% 20%, rgba(147, 197, 253, 0.24), transparent 36%),
            linear-gradient(125deg, var(--blue-900) 0%, var(--blue-700) 100%);
        border-radius: 24px;
        padding: 28px 32px;
        color: #ffffff;
        margin-top: 34px;
        margin-bottom: 18px;
        box-shadow: var(--shadow-lg);
        text-align: {text_align};
    }}

    .report-banner small {{
        display: block;
        color: #bfdbfe;
        font-size: 11px;
        font-weight: 800;
        letter-spacing: 1.2px;
        margin-bottom: 6px;
    }}

    .report-banner h2 {{
        color: #ffffff;
        font-size: 26px;
        font-weight: 900;
        margin: 0;
    }}

    .score-panel {{
        background: var(--surface);
        border: 1px solid var(--border);
        border-radius: 22px;
        padding: 24px 18px;
        text-align: center;
        box-shadow: var(--shadow-sm);
        height: 100%;
    }}

    .score-ring {{
        width: 156px;
        height: 156px;
        border-radius: 50%;
        margin: 16px auto;
        display: flex;
        align-items: center;
        justify-content: center;
        position: relative;
        box-shadow: inset 0 0 0 1px var(--border);
    }}

    .score-ring::before {{
        content: "";
        position: absolute;
        width: 126px;
        height: 126px;
        background: var(--surface);
        border-radius: 50%;
        box-shadow: var(--shadow-sm);
    }}

    .score-num {{
        position: relative;
        z-index: 2;
        font-size: 36px;
        font-weight: 900;
        color: var(--blue-800);
    }}

    .score-title {{
        font-size: 14px;
        font-weight: 800;
        color: var(--muted);
    }}

    .score-badge {{
        display: inline-block;
        background: var(--blue-50);
        color: var(--blue-700);
        border: 1px solid var(--blue-200);
        padding: 6px 16px;
        border-radius: 999px;
        font-size: 12.5px;
        font-weight: 800;
    }}

    .insight-card {{
        background: var(--surface);
        border: 1px solid var(--border);
        border-radius: 22px;
        padding: 22px 24px;
        box-shadow: var(--shadow-sm);
        transition: border-color 0.2s ease, box-shadow 0.2s ease;
        height: 100%;
    }}

    .insight-card:hover {{
        border-color: var(--blue-200);
        box-shadow: var(--shadow-md);
    }}

    .insight-head {{
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 10px;
        padding-bottom: 12px;
        margin-bottom: 12px;
        border-bottom: 1px solid #eef3fa;
    }}

    .insight-title {{
        display: flex;
        align-items: center;
        gap: 8px;
        font-size: 15.5px;
        font-weight: 900;
        color: var(--blue-900);
    }}

    .insight-pill {{
        background: var(--blue-50);
        color: var(--blue-700);
        font-size: 11.5px;
        font-weight: 800;
        padding: 3px 10px;
        border-radius: 999px;
    }}

    .insight-row {{
        color: #475569;
        font-size: 13.5px;
        line-height: 1.85;
        padding: 8px 0;
        border-bottom: 1px dashed #e8eef7;
        text-align: {text_align};
        overflow-wrap: anywhere;
    }}

    .insight-row:last-child {{
        border-bottom: none;
        padding-bottom: 2px;
    }}

    .disclaimer-text {{
        text-align: center;
        color: var(--muted);
        font-size: 12px;
        margin-top: 18px;
        line-height: 1.8;
    }}

    .saas-footer {{
        margin-top: 52px;
        padding-top: 24px;
        border-top: 1px solid var(--border);
        text-align: center;
        color: var(--muted);
        font-size: 12.5px;
    }}

    .saas-footer strong {{
        color: var(--blue-800);
        font-size: 16px;
        font-weight: 900;
    }}

    @media (max-width: 768px) {{
        .hero-box {{
            padding: 38px 20px;
            border-radius: 22px;
        }}
        .upload-shell {{
            padding: 20px 16px 14px;
            border-radius: 20px;
        }}
        .report-banner {{
            padding: 22px 18px;
            border-radius: 20px;
        }}
    }}
    </style>
    """,
    unsafe_allow_html=True,
)


# =========================
# TOP BAR & LANGUAGE BUTTON
# =========================

st.markdown('<div class="topbar-card">', unsafe_allow_html=True)
top_left, top_right = st.columns([8, 2], vertical_alignment="center")

with top_left:
    st.markdown(
        """
        <div class="brand-inline">
            <div class="brand-badge">✦</div>
            <div>
                <div class="brand-title">CVision <span>AI</span></div>
                <div class="brand-caption">Smart Resume &amp; ATS Evaluation Suite</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with top_right:
    st.markdown('<div class="lang-btn-box">', unsafe_allow_html=True)
    st.button(
        t["switch_label"],
        key="language_toggle_btn",
        help=t["switch_help"],
        on_click=toggle_language,
    )
    st.markdown("</div>", unsafe_allow_html=True)

st.markdown("</div>", unsafe_allow_html=True)


# =========================
# HERO SECTION
# =========================

st.markdown(
    f"""
    <div class="hero-box">
        <div class="hero-ring"></div>
        <div class="hero-inner">
            <div class="hero-pill">✦ {html.escape(t["badge"])}</div>
            <h1>{html.escape(t["hero"])}<br><span>{html.escape(t["hero_highlight"])}</span></h1>
            <p>{html.escape(t["description"])}</p>
            <div class="hero-trust">✧ {html.escape(t["secure"])}</div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)


# =========================
# 3-STEP OVERVIEW
# =========================

st.markdown(f'<div class="kicker">{html.escape(t["step"])}</div>', unsafe_allow_html=True)

steps = [
    ("📄", "f1", "f1d"),
    ("✦", "f2", "f2d"),
    ("🚀", "f3", "f3d"),
]

step_cols = st.columns(3)
for col, (icon, title_k, desc_k) in zip(step_cols, steps):
    with col:
        st.markdown(
            f"""
            <div class="step-card">
                <div class="step-icon">{icon}</div>
                <strong>{html.escape(t[title_k])}</strong>
                <span>{html.escape(t[desc_k])}</span>
            </div>
            """,
            unsafe_allow_html=True,
        )


# =========================
# UPLOAD & ANALYSIS WORKFLOW
# =========================

st.markdown(
    f"""
    <div class="upload-shell">
        <div class="upload-title">📄 {html.escape(t["upload"])}</div>
        <div class="upload-sub">{html.escape(t["upload_hint"])}</div>
    """,
    unsafe_allow_html=True,
)

uploaded_file = st.file_uploader(
    t["drop"],
    type=["pdf", "docx"],
    key="cv_file_uploader",
    help=t["formats"],
)

st.markdown(
    f'<div style="color:#64748b;font-size:12px;margin-top:6px;text-align:{text_align};">{html.escape(t["formats"])}</div>',
    unsafe_allow_html=True,
)

if uploaded_file is not None:
    size_mb = uploaded_file.size / (1024 * 1024)
    if uploaded_file.size > 10 * 1024 * 1024:
        st.error(t["error_size"])
        st.stop()

    st.markdown(
        f"""
        <div class="file-meta-card">
            <div class="file-meta-main">
                <span>✓</span>
                <span>{html.escape(t["file_ready"])}: {html.escape(uploaded_file.name)}</span>
            </div>
            <div class="file-meta-tag">{html.escape(t["file_size_label"])}: {size_mb:.2f} MB</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown('<div class="action-btn-wrap">', unsafe_allow_html=True)
    if st.button(t["analyze"], key="run_analysis_btn", use_container_width=True):
        try:
            with st.spinner(t["loading"]):
                extracted = extract_text(uploaded_file.getvalue(), uploaded_file.name)
                if not extracted or len(extracted.strip()) < 15:
                    st.error(t["error_text"])
                    st.stop()

                analysis_data = analyze_cv(extracted, lang)
                st.session_state.cv_text = extracted
                st.session_state.file_name = uploaded_file.name
                st.session_state.result = analysis_data
                st.session_state.result_lang = lang

            st.rerun()
        except KeyError:
            st.error(t["error_key"])
        except Exception:
            st.error(t["error_api"])
    st.markdown("</div>", unsafe_allow_html=True)

st.markdown("</div>", unsafe_allow_html=True)


# =========================
# RESULTS SECTION & FOOTER
# =========================

def _render_insight_card(icon: str, title: str, items: list, empty_label: str, items_label: str) -> str:
    clean_items = [str(item).strip() for item in (items or []) if str(item).strip()]
    if not clean_items:
        rows_html = f'<div class="insight-row">{html.escape(empty_label)}</div>'
    else:
        rows_html = "".join(
            f'<div class="insight-row">• {html.escape(entry)}</div>' for entry in clean_items
        )

    return (
        f'<div class="insight-card">'
        f'  <div class="insight-head">'
        f'    <div class="insight-title"><span>{icon}</span><span>{html.escape(title)}</span></div>'
        f'    <span class="insight-pill">{len(clean_items)} {html.escape(items_label)}</span>'
        f'  </div>'
        f'  <div>{rows_html}</div>'
        f'</div>'
    )


if st.session_state.result is not None:
    res = st.session_state.result
    raw_score = res.get("compatibilityScore", 0)
    try:
        score = max(0, min(100, int(round(float(raw_score)))))
    except (TypeError, ValueError):
        score = 0

    if score >= 80:
        verdict = t["excellent"]
        ring_color = "#1d4ed8"
    elif score >= 60:
        verdict = t["good"]
        ring_color = "#2563eb"
    elif score >= 40:
        verdict = t["average"]
        ring_color = "#3b82f6"
    else:
        verdict = t["weak"]
        ring_color = "#64748b"

    deg = int((score / 100) * 360)
    file_caption = (
        f" — {st.session_state.file_name}" if st.session_state.file_name else ""
    )

    st.markdown(
        f"""
        <div class="report-banner">
            <small>{html.escape(t["result_label"])}{html.escape(file_caption)}</small>
            <h2>{html.escape(t["result_title"])}</h2>
        </div>
        """,
        unsafe_allow_html=True,
    )

    left_res, right_res = st.columns([4, 8])

    with left_res:
        st.markdown(
            f"""
            <div class="score-panel">
                <div class="score-title">{html.escape(t["score"])}</div>
                <div class="score-ring" style="background: conic-gradient({ring_color} {deg}deg, #e2ebf7 {deg}deg 360deg);">
                    <div class="score-num">{score}%</div>
                </div>
                <div class="score-badge">{html.escape(verdict)}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with right_res:
        summary_text = str(res.get("summary") or "").strip() or t["empty"]
        st.markdown(
            f"""
            <div class="insight-card">
                <div class="insight-head">
                    <div class="insight-title"><span>✦</span><span>{html.escape(t["summary"])}</span></div>
                </div>
                <div class="insight-row">{html.escape(summary_text)}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown("<div style='height:18px;'></div>", unsafe_allow_html=True)

    grid_top_1, grid_top_2 = st.columns(2)
    with grid_top_1:
        st.markdown(
            _render_insight_card(
                "✓",
                t["skills"],
                res.get("matchingSkills", []),
                t["empty"],
                t["items_count"],
            ),
            unsafe_allow_html=True,
        )
    with grid_top_2:
        st.markdown(
            _render_insight_card(
                "△",
                t["missing"],
                res.get("missingSkills", []),
                t["empty"],
                t["items_count"],
            ),
            unsafe_allow_html=True,
        )

    st.markdown("<div style='height:18px;'></div>", unsafe_allow_html=True)

    grid_bot_1, grid_bot_2 = st.columns(2)
    with grid_bot_1:
        st.markdown(
            _render_insight_card(
                "★",
                t["strengths"],
                res.get("strengths", []),
                t["empty"],
                t["items_count"],
            ),
            unsafe_allow_html=True,
        )
    with grid_bot_2:
        st.markdown(
            _render_insight_card(
                "↗",
                t["suggestions"],
                res.get("suggestions", []),
                t["empty"],
                t["items_count"],
            ),
            unsafe_allow_html=True,
        )

    st.markdown(
        f'<div class="disclaimer-text">{html.escape(t["analysis_note"])}</div>',
        unsafe_allow_html=True,
    )

    st.markdown("<div style='height:18px;'></div>", unsafe_allow_html=True)
    st.markdown('<div class="action-btn-wrap">', unsafe_allow_html=True)
    if st.button(t["new"], key="reset_analysis_btn", use_container_width=True):
        st.session_state.result = None
        st.session_state.result_lang = None
        st.session_state.cv_text = ""
        st.session_state.file_name = ""
        st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)

st.markdown(
    f"""
    <div class="saas-footer">
        <strong>CVision AI</strong>
        <div style="margin-top:6px;">{html.escape(t["footer"])}</div>
    </div>
    """,
    unsafe_allow_html=True,
)
