
import streamlit as st
import json
import os
import re
import html
from io import BytesIO
from pypdf import PdfReader
from docx import Document
from google import genai
from google.genai import types

st.set_page_config(
    page_title="CVision AI",
    page_icon="📄",
    layout="wide",
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

def toggle_language():
    st.session_state.lang = (
        "en" if st.session_state.lang == "ar" else "ar"
    )

# =========================
# LANGUAGE SWITCH
# =========================

with st.container(key="language_bar"):
    space, lang_col = st.columns([10, 1])

    with lang_col:
        st.button(
            "EN" if st.session_state.lang == "ar" else "عربي",
            key="language_toggle",
            on_click=toggle_language,
            use_container_width=True,
            type="secondary"
        )

# =========================
# TRANSLATIONS
# =========================

TEXT = {
    "ar": {
        "subtitle": "منصة ذكية لتحليل وتقييم السير الذاتية",
        "badge": "ذكاء اصطناعي • تحليل احترافي",
        "hero": "سيرتك الذاتية،",
        "hero_highlight": "بشكل أقوى.",
        "description": "اكتشف نقاط قوتك، قيّم مهاراتك، واعرف كيف تطوّر سيرتك الذاتية لتكون أكثر احترافية وجاهزية لسوق العمل.",
        "upload": "حلّل سيرتك الذاتية",
        "upload_hint": "ارفع ملفك ودعنا نساعدك على اكتشاف إمكانياته.",
        "drop": "اسحب ملفك هنا أو اضغط للاختيار",
        "formats": "PDF أو DOCX • حتى 10 ميجابايت",
        "analyze": "ابدأ التحليل الذكي",
        "loading": "نحلل سيرتك الذاتية الآن...",
        "score": "درجة جاهزية السيرة الذاتية",
        "summary": "نظرة عامة",
        "skills": "المهارات المكتشفة",
        "missing": "عناصر تحتاج إلى تحسين",
        "strengths": "نقاط القوة",
        "suggestions": "توصياتنا لك",
        "new": "تحليل سيرة ذاتية جديدة",
        "error_size": "حجم الملف أكبر من 10 ميجابايت.",
        "error_text": "تعذر استخراج النص من الملف. تأكد أن الملف يحتوي على نص قابل للقراءة.",
        "error_api": "تعذر إتمام التحليل. تأكد من إعدادات الخدمة وحاول مرة أخرى.",
        "footer": "خطوة أذكى نحو فرصتك المهنية القادمة.",
        "f1": "ارفع ملفك",
        "f1d": "ارفع سيرتك الذاتية بسهولة.",
        "f2": "تحليل ذكي",
        "f2d": "اكتشف مهاراتك وخبراتك.",
        "f3": "طوّر سيرتك",
        "f3d": "احصل على توصيات عملية.",
        "step": "ثلاث خطوات بسيطة",
        "result_label": "نتيجة التحليل",
        "excellent": "ممتاز",
        "good": "جيد جدًا",
        "average": "يحتاج إلى تطوير",
        "weak": "يحتاج إلى تحسين",
        "empty": "لا توجد عناصر لعرضها حاليًا.",
        "file_ready": "ملفك جاهز للتحليل",
        "secure": "تحليل سريع • تجربة سهلة • نتائج واضحة",
        "analysis_note": "هذا التقييم استرشادي لجودة السيرة الذاتية، وليس ضمانًا للقبول الوظيفي."
    },

    "en": {
        "subtitle": "Your intelligent CV analysis platform",
        "badge": "ARTIFICIAL INTELLIGENCE • PROFESSIONAL ANALYSIS",
        "hero": "Make your CV",
        "hero_highlight": "stand out.",
        "description": "Discover your strengths, evaluate your skills, and learn how to improve your resume to become more professional and job-ready.",
        "upload": "Analyze Your Resume",
        "upload_hint": "Upload your file and let us uncover its potential.",
        "drop": "Drag your file here or click to browse",
        "formats": "PDF or DOCX • Up to 10 MB",
        "analyze": "Start Smart Analysis",
        "loading": "Analyzing your CV...",
        "score": "CV Readiness Score",
        "summary": "Overview",
        "skills": "Identified Skills",
        "missing": "Areas to Improve",
        "strengths": "Your Strengths",
        "suggestions": "Our Recommendations",
        "new": "Analyze Another CV",
        "error_size": "File size exceeds 10 MB.",
        "error_text": "Could not extract readable text from this file.",
        "error_api": "Analysis failed. Please check the service configuration and try again.",
        "footer": "A smarter step towards your next career opportunity.",
        "f1": "Upload your CV",
        "f1d": "Upload your resume easily.",
        "f2": "Smart analysis",
        "f2d": "Discover your skills and experience.",
        "f3": "Improve your CV",
        "f3d": "Get practical recommendations.",
        "step": "THREE SIMPLE STEPS",
        "result_label": "ANALYSIS RESULT",
        "excellent": "Excellent",
        "good": "Very Good",
        "average": "Needs Development",
        "weak": "Needs Improvement",
        "empty": "No items to display.",
        "file_ready": "Your file is ready for analysis",
        "secure": "Fast analysis • Easy experience • Clear results",
        "analysis_note": "This score is an indicative CV quality assessment, not a guarantee of employment."
    }
}

# =========================
# CURRENT LANGUAGE
# =========================

ar = st.session_state.lang == "ar"
t = TEXT[st.session_state.lang]
font = "Cairo" if ar else "Inter"
direction = "rtl" if ar else "ltr"
text_align = "right" if ar else "left"

# =========================
# DESIGN SYSTEM
# =========================

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;500;600;700;800;900&family=Inter:wght@400;500;600;700;800;900&display=swap');

:root {{
    --primary: #2563eb;
    --primary-dark: #1745b8;
    --text: #10213f;
    --muted: #718096;
    --border: #e6edf7;
    --surface: #ffffff;
}}

* {{
    box-sizing: border-box;
}}

.stApp {{
    background:
        radial-gradient(ellipse at 5% 5%, rgba(59,130,246,.13), transparent 32%),
        radial-gradient(ellipse at 95% 20%, rgba(96,165,250,.13), transparent 30%),
        linear-gradient(180deg, #f8faff 0%, #f2f6fd 100%);
    color: var(--text);
    font-family: '{font}', sans-serif;
    direction: {direction};
}}

html, body, [class*="css"] {{
    font-family: '{font}', sans-serif;
}}

[data-testid="stAppViewContainer"] {{
    direction: {direction};
}}

[data-testid="stHeader"] {{
    background: transparent;
}}

.block-container {{
    max-width: 1180px;
    padding-top: 1rem;
    padding-bottom: 2rem;
}}

h1, h2, h3, p {{
    font-family: '{font}', sans-serif;
}}

#MainMenu, footer {{
    visibility: hidden;
}}

/* LANGUAGE BUTTON */

.st-key-language_bar {{
    display: flex;
    justify-content: flex-end;
    margin-bottom: 5px;
}}

.st-key-language_bar div.stButton > button {{
    min-height: 36px;
    width: 65px;
    border-radius: 11px;
    background: white;
    color: #2563eb;
    border: 1px solid #dbeafe;
    font-size: 12px;
    font-weight: 900;
    padding: 0 12px;
    box-shadow: 0 4px 13px rgba(37,99,235,.08);
}}

.st-key-language_bar div.stButton > button:hover {{
    background: #eff6ff;
    color: #1745b8;
    border-color: #93c5fd;
}}

/* BRAND */

.brand {{
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 12px;
    margin: 8px 0 0;
}}

.brand-icon {{
    width: 48px;
    height: 48px;
    display: flex;
    align-items: center;
    justify-content: center;
    background: linear-gradient(135deg,#2563eb,#60a5fa);
    color: white;
    border-radius: 16px;
    font-size: 23px;
    box-shadow: 0 9px 25px rgba(37,99,235,.25);
}}

.brand-name {{
    font-size: 30px;
    font-weight: 900;
    letter-spacing: -1px;
    color: #173b79;
}}

.brand-name span {{
    color: #3b82f6;
}}

.brand-sub {{
    color: #8493aa;
    font-size: 12px;
    text-align: center;
    margin-top: 5px;
}}

/* HERO */

.hero {{
    position: relative;
    overflow: hidden;
    background:
        radial-gradient(circle at 85% 15%, rgba(147,197,253,.22), transparent 28%),
        linear-gradient(120deg,#102d61 0%,#174b9d 55%,#347ce5 100%);
    border: 1px solid rgba(255,255,255,.12);
    border-radius: 30px;
    padding: 76px 35px;
    margin: 30px 0 32px;
    text-align: center;
    box-shadow: 0 24px 60px rgba(23,69,184,.20);
}}

.hero:before {{
    content: "";
    position: absolute;
    width: 370px;
    height: 370px;
    border-radius: 50%;
    border: 1px solid rgba(255,255,255,.16);
    top: -190px;
    left: -65px;
}}

.hero:after {{
    content: "";
    position: absolute;
    width: 250px;
    height: 250px;
    border-radius: 50%;
    background: rgba(255,255,255,.07);
    bottom: -145px;
    right: 5%;
}}

.hero-content {{
    position: relative;
    z-index: 2;
}}

.hero-badge {{
    display: inline-block;
    color: #dbeafe;
    background: rgba(255,255,255,.12);
    border: 1px solid rgba(255,255,255,.20);
    border-radius: 50px;
    padding: 10px 19px;
    font-size: 12px;
    font-weight: 800;
    margin-bottom: 22px;
    backdrop-filter: blur(8px);
}}

.hero h1 {{
    color: white;
    font-size: clamp(35px,5vw,58px);
    line-height: 1.35;
    font-weight: 900;
    margin: 0 0 17px;
}}

.hero h1 span {{
    color: #93c5fd;
}}

.hero p {{
    color: #dbeafe;
    font-size: 15px;
    line-height: 2;
    max-width: 700px;
    margin: 0 auto;
}}

.hero-note {{
    display: inline-block;
    color: #dbeafe;
    font-size: 11px;
    margin-top: 28px;
    opacity: .9;
    background: rgba(255,255,255,.08);
    padding: 9px 17px;
    border-radius: 30px;
}}

/* SECTION HEADINGS */

.section-kicker {{
    color: #3b82f6;
    font-size: 11px;
    font-weight: 900;
    letter-spacing: 1.4px;
    text-align: center;
    margin-top: 35px;
    margin-bottom: 18px;
}}

/* FEATURE CARDS */

.feature-card {{
    background: rgba(255,255,255,.94);
    border: 1px solid #e7edf7;
    border-radius: 21px;
    padding: 25px 17px;
    min-height: 160px;
    box-shadow: 0 8px 24px rgba(31,65,114,.05);
    transition: all .25s ease;
    text-align: center;
}}

.feature-card:hover {{
    transform: translateY(-5px);
    border-color: #bfdbfe;
    box-shadow: 0 16px 34px rgba(37,99,235,.11);
}}

.feature-icon {{
    width: 48px;
    height: 48px;
    border-radius: 15px;
    display: flex;
    align-items: center;
    justify-content: center;
    background: linear-gradient(135deg,#eff6ff,#dbeafe);
    font-size: 21px;
    margin: 0 auto 14px;
}}

.feature-card strong {{
    display: block;
    color: #173b79;
    font-size: 15px;
    margin-bottom: 7px;
}}

.feature-card span {{
    color: #8190a8;
    font-size: 12px;
    line-height: 1.8;
}}

/* UPLOAD PANEL */

.upload-panel {{
    background: rgba(255,255,255,.97);
    border: 1px solid #e3ebf7;
    border-radius: 25px;
    padding: 30px;
    margin-top: 36px;
    margin-bottom: 14px;
    box-shadow: 0 15px 40px rgba(31,65,114,.07);
}}

.panel-heading {{
    color: #15366c;
    font-size: 23px;
    font-weight: 900;
    text-align: center;
    margin-bottom: 7px;
}}

.panel-description {{
    color: #8290a7;
    font-size: 13px;
    text-align: center;
    margin-bottom: 5px;
}}

[data-testid="stFileUploader"] {{
    border: 2px dashed #bfdbfe;
    border-radius: 19px;
    background: #f7faff;
    padding: 16px;
    transition: all .2s ease;
}}

[data-testid="stFileUploader"]:hover {{
    border-color: #3b82f6;
    background: #eff6ff;
}}

[data-testid="stFileUploader"] section {{
    border: none;
}}

[data-testid="stFileUploader"] button {{
    border-radius: 10px;
}}

/* GENERAL BUTTONS */

div.stButton > button {{
    background: linear-gradient(110deg,#2563eb,#3478e5);
    color: white;
    border: none;
    border-radius: 13px;
    min-height: 49px;
    font-family: '{font}', sans-serif;
    font-weight: 800;
    box-shadow: 0 8px 20px rgba(37,99,235,.18);
    transition: all .2s ease;
}}

div.stButton > button:hover {{
    background: linear-gradient(110deg,#1745b8,#2563eb);
    color: white;
    border: none;
    transform: translateY(-2px);
    box-shadow: 0 12px 25px rgba(37,99,235,.25);
}}

[data-testid="stProgress"] > div > div {{
    background: linear-gradient(90deg,#2563eb,#60a5fa);
    border-radius: 20px;
}}

/* SUCCESS AND ALERTS */

[data-testid="stAlert"] {{
    border-radius: 14px;
    border: 1px solid #e3ebf7;
}}

/* RESULTS */

.result-header {{
    background:
        radial-gradient(circle at 90% 10%, rgba(147,197,253,.20), transparent 30%),
        linear-gradient(120deg,#102d61,#2563eb);
    border-radius: 23px;
    padding: 30px;
    color: white;
    margin-top: 32px;
    text-align: center;
    box-shadow: 0 14px 35px rgba(37,99,235,.16);
}}

.result-header small {{
    color: #bfdbfe;
    letter-spacing: 1.5px;
    font-weight: 800;
}}

.result-header h2 {{
    color: white;
    font-size: 27px;
    font-weight: 900;
    margin: 9px 0 0;
}}

.score-card {{
    background: white;
    border: 1px solid #e5edf8;
    border-radius: 22px;
    padding: 27px 15px;
    text-align: center;
    box-shadow: 0 12px 30px rgba(31,65,114,.06);
    height: 100%;
}}

.score-circle {{
    width: 155px;
    height: 155px;
    border-radius: 50%;
    margin: 18px auto;
    display: flex;
    align-items: center;
    justify-content: center;
    position: relative;
}}

.score-circle:before {{
    content: "";
    position: absolute;
    width: 127px;
    height: 127px;
    background: white;
    border-radius: 50%;
}}

.score-value {{
    position: relative;
    z-index: 1;
    color: #1745a0;
    font-size: 38px;
    font-weight: 900;
}}

.score-label {{
    color: #74849d;
    font-size: 13px;
    font-weight: 800;
}}

.score-status {{
    display: inline-block;
    background: #eff6ff;
    color: #2563eb;
    border-radius: 30px;
    padding: 8px 18px;
    font-size: 12px;
    font-weight: 800;
}}

.result-card {{
    background: white;
    border: 1px solid #e7edf7;
    border-radius: 19px;
    padding: 22px;
    min-height: 160px;
    margin-bottom: 15px;
    box-shadow: 0 6px 22px rgba(31,65,114,.045);
    transition: all .2s ease;
}}

.result-card:hover {{
    border-color: #d5e5fb;
    box-shadow: 0 11px 27px rgba(37,99,235,.08);
}}

.result-card-title {{
    color: #183c78;
    font-weight: 900;
    font-size: 15px;
    margin-bottom: 13px;
    text-align: center;
}}

.result-item {{
    color: #5b6b82;
    font-size: 13px;
    line-height: 1.9;
    padding: 8px 0;
    border-bottom: 1px solid #f0f3f8;
    text-align: {text_align};
    overflow-wrap: anywhere;
}}

.result-item:last-child {{
    border-bottom: none;
}}

.footer {{
    border-top: 1px solid #e8eef7;
    padding-top: 25px;
    margin-top: 55px;
    text-align: center;
    color: #8795aa;
    font-size: 12px;
}}

.footer strong {{
    color: #2563eb;
    font-size: 17px;
}}

.analysis-note {{
    color: #8492a8;
    font-size: 11px;
    text-align: center;
    margin-top: 15px;
    line-height: 1.8;
}}

/* MOBILE */

@media (max-width: 700px) {{
    .block-container {{
        padding-left: 1rem;
        padding-right: 1rem;
    }}

    .hero {{
        padding: 45px 18px;
        border-radius: 22px;
    }}

    .hero h1 {{
        font-size: 34px;
    }}

    .upload-panel {{
        padding: 19px;
    }}

    .brand-name {{
        font-size: 25px;
    }}

    .feature-card {{
        padding: 17px 10px;
        min-height: 155px;
    }}

    .result-card {{
        padding: 16px;
    }}

    .result-header {{
        padding: 23px 15px;
    }}
}}
</style>
""", unsafe_allow_html=True)

# =========================
# BRAND HEADER
# =========================

st.markdown("""
<div class="brand">
    <div class="brand-icon">✦</div>
    <div class="brand-name">CVision <span>AI</span></div>
</div>
""", unsafe_allow_html=True)

st.markdown(
    f'<div class="brand-sub">{t["subtitle"]}</div>',
    unsafe_allow_html=True
)

# =========================
# HERO
# =========================

st.markdown(f"""
<div class="hero">
    <div class="hero-content">
        <div class="hero-badge">✦ {t["badge"]}</div>
        <h1>{t["hero"]}<br><span>{t["hero_highlight"]}</span></h1>
        <p>{t["description"]}</p>
        <div class="hero-note">✧ {t["secure"]}</div>
    </div>
</div>
""", unsafe_allow_html=True)

# =========================
# FEATURES
# =========================

st.markdown(
    f'<div class="section-kicker">{t["step"]}</div>',
    unsafe_allow_html=True
)

features = [
    ("📄", "f1", "f1d"),
    ("✦", "f2", "f2d"),
    ("🚀", "f3", "f3d")
]

cols = st.columns(3)

for col, (icon, title_key, desc_key) in zip(cols, features):
    with col:
        st.markdown(f"""
        <div class="feature-card">
            <div class="feature-icon">{icon}</div>
            <strong>{t[title_key]}</strong>
            <span>{t[desc_key]}</span>
        </div>
        """, unsafe_allow_html=True)

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

    except Exception as e:
        raise RuntimeError("Analysis service unavailable") from e


# =========================
# UPLOAD SECTION
# =========================

st.markdown(f"""
<div class="upload-panel">
    <div class="panel-heading">📄 {t["upload"]}</div>
    <div class="panel-description">{t["upload_hint"]}</div>
</div>
""", unsafe_allow_html=True)

uploaded_file = st.file_uploader(
    t["drop"],
    type=["pdf", "docx"],
    key="cv_upload",
    help=t["formats"]
)

st.markdown(
    f'<div style="text-align:center;color:#8492a8;font-size:11px;margin-top:-5px;">{t["formats"]}</div>',
    unsafe_allow_html=True
)

if uploaded_file:

    if uploaded_file.size > 10 * 1024 * 1024:
        st.error(t["error_size"])
        st.stop()

    st.success(f"✓ {t['file_ready']}: {uploaded_file.name}")

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

# =========================
# RESULTS
# =========================

if st.session_state.result:

    result = st.session_state.result

    score = max(
        0,
        min(100, int(result.get("compatibilityScore", 0)))
    )

    if score >= 85:
        status = t["excellent"]
    elif score >= 70:
        status = t["good"]
    elif score >= 50:
        status = t["average"]
    else:
        status = t["weak"]

    st.markdown(f"""
    <div class="result-header">
        <small>{t["result_label"]}</small>
        <h2>📊 {t["summary"]}</h2>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("<div style='height:18px'></div>", unsafe_allow_html=True)

    score_col, summary_col = st.columns([1, 1.5])

    with score_col:
        st.markdown(f"""
        <div class="score-card">
            <div class="score-label">{t["score"]}</div>
            <div class="score-circle"
                 style="background:conic-gradient(#3b82f6 {score}%,#e8effa 0);">
                <div class="score-value">{score}%</div>
            </div>
            <div class="score-status">{status}</div>
        </div>
        """, unsafe_allow_html=True)

    with summary_col:
        safe_summary = html.escape(str(result.get("summary", "")))

        st.markdown(f"""
        <div class="result-card" style="height:100%;">
            <div class="result-card-title">✦ {t["summary"]}</div>
            <div style="color:#5b6b82;font-size:14px;line-height:2;text-align:center;">
                {safe_summary}
            </div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)

    sections = [
        ("💡", "skills", "matchingSkills"),
        ("⚠️", "missing", "missingSkills"),
        ("⭐", "strengths", "strengths"),
        ("🚀", "suggestions", "suggestions")
    ]

    result_cols = st.columns(2)

    for index, (icon, label, key) in enumerate(sections):

        with result_cols[index % 2]:

            items = result.get(key, [])

            if not isinstance(items, list):
                items = []

            content = ""

            for item in items:
                safe_item = html.escape(str(item))
                content += f'<div class="result-item">• {safe_item}</div>'

            if not content:
                content = f'<div class="result-item">{t["empty"]}</div>'

            st.markdown(f"""
            <div class="result-card">
                <div class="result-card-title">{icon} {t[label]}</div>
                {content}
            </div>
            """, unsafe_allow_html=True)

    st.markdown(
        f'<div class="analysis-note">{t["analysis_note"]}</div>',
        unsafe_allow_html=True
    )

    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)

    if st.button(t["new"], use_container_width=True):
        st.session_state.result = None
        st.session_state.cv_text = ""
        st.rerun()

# =========================
# FOOTER
# =========================

st.markdown(f"""
<div class="footer">
    <strong>CVision AI</strong>
    <div style="margin-top:6px;">{t["footer"]}</div>
    <div style="margin-top:12px;font-size:10px;">© 2026 CVision AI. All rights reserved.</div>
</div>
""", unsafe_allow_html=True)
