
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

t = TEXT[st.session_state.lang]
ar = st.session_state.lang == "ar"
font = "Cairo" if ar else "Inter"

# =========================
# DESIGN
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
}}

.stApp {{
    background:
        radial-gradient(ellipse at 10% 5%, rgba(59,130,246,.12), transparent 30%),
        radial-gradient(ellipse at 90% 25%, rgba(96,165,250,.12), transparent 28%),
        #f7f9fd;
    color: var(--text);
    font-family: '{font}', sans-serif;
    direction: {'rtl' if ar else 'ltr'};
}}

html, body, [class*="css"] {{
    font-family: '{font}', sans-serif;
}}

[data-testid="stHeader"] {{
    background: transparent;
}}

.block-container {{
    max-width: 1160px;
    padding-top: 1rem;
    padding-bottom: 2rem;
}}

h1, h2, h3, p {{
    font-family: '{font}', sans-serif;
}}

h1, h2, h3 {{
    color: var(--text);
}}

#MainMenu, footer {{
    visibility: hidden;
}}

/* Language button */

div[data-testid="stHorizontalBlock"]:first-of-type
div[data-testid="stButton"] button {{
    min-height: 36px;
    border-radius: 11px;
    background: white;
    color: #2563eb;
    border: 1px solid #dbeafe;
    font-size: 12px;
    font-weight: 900;
    padding: 0 12px;
    box-shadow: 0 3px 12px rgba(37,99,235,.07);
}}

div[data-testid="stHorizontalBlock"]:first-of-type
div[data-testid="stButton"] button:hover {{
    background: #eff6ff;
    color: #1745b8;
    border-color: #93c5fd;
    transform: translateY(-1px);
}}

/* Brand */

.brand {{
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 10px;
    margin: 5px 0 0;
}}

.brand-icon {{
    width: 44px;
    height: 44px;
    display: flex;
    align-items: center;
    justify-content: center;
    background: linear-gradient(135deg,#2563eb,#60a5fa);
    color: white;
    border-radius: 15px;
    font-size: 22px;
    box-shadow: 0 8px 22px rgba(37,99,235,.27);
}}

.brand-name {{
    font-size: 28px;
    font-weight: 900;
    letter-spacing: -1px;
    color: #173b79;
}}

.brand-name span {{
    color: #3b82f6;
}}

.brand-sub {{
    color: #8493aa;
    font-size: 11px;
    letter-spacing: 1.2px;
    text-align: center;
    margin-top: 3px;
}}

/* Hero */

.hero {{
    position: relative;
    overflow: hidden;
    background: linear-gradient(120deg,#102d61 0%,#174b9d 55%,#347ce5 100%);
    border-radius: 30px;
    padding: 70px 35px;
    margin: 30px 0 28px;
    text-align: center;
    box-shadow: 0 22px 55px rgba(23,69,184,.20);
}}

.hero:before {{
    content: "";
    position: absolute;
    width: 330px;
    height: 330px;
    border-radius: 50%;
    border: 1px solid rgba(255,255,255,.15);
    top: -170px;
    left: -60px;
}}

.hero:after {{
    content: "";
    position: absolute;
    width: 230px;
    height: 230px;
    border-radius: 50%;
    background: rgba(255,255,255,.07);
    bottom: -130px;
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
    padding: 9px 18px;
    font-size: 12px;
    font-weight: 700;
    margin-bottom: 20px;
    backdrop-filter: blur(8px);
}}

.hero h1 {{
    color: white;
    font-size: clamp(32px,5vw,55px);
    line-height: 1.35;
    font-weight: 900;
    margin: 0 0 14px;
}}

.hero h1 span {{
    color: #93c5fd;
}}

.hero p {{
    color: #dbeafe;
    font-size: 15px;
    line-height: 2;
    max-width: 680px;
    margin: 0 auto;
}}

.hero-note {{
    color: #dbeafe;
    font-size: 11px;
    margin-top: 25px;
    opacity: .85;
}}

.section-kicker {{
    color: #3b82f6;
    font-size: 11px;
    font-weight: 800;
    letter-spacing: 1.4px;
    text-align: center;
    margin-top: 34px;
}}

.section-heading {{
    font-size: 25px;
    font-weight: 900;
    color: #142b50;
    text-align: center;
    margin: 5px 0 22px;
}}

/* Feature cards */

.feature-card {{
    background: rgba(255,255,255,.92);
    border: 1px solid #e7edf7;
    border-radius: 20px;
    padding: 23px 17px;
    min-height: 145px;
    box-shadow: 0 7px 22px rgba(31,65,114,.045);
    transition: all .25s ease;
    text-align: center;
}}

.feature-card:hover {{
    transform: translateY(-5px);
    border-color: #bfdbfe;
    box-shadow: 0 15px 30px rgba(37,99,235,.10);
}}

.feature-icon {{
    width: 44px;
    height: 44px;
    border-radius: 14px;
    display: flex;
    align-items: center;
    justify-content: center;
    background: linear-gradient(135deg,#eff6ff,#dbeafe);
    font-size: 20px;
    margin: 0 auto 13px;
}}

.feature-card strong {{
    display: block;
    color: #173b79;
    font-size: 15px;
    margin-bottom: 6px;
}}

.feature-card span {{
    color: #8190a8;
    font-size: 12px;
    line-height: 1.8;
}}

/* Upload */

.upload-panel {{
    background: rgba(255,255,255,.96);
    border: 1px solid #e3ebf7;
    border-radius: 25px;
    padding: 30px;
    margin-top: 35px;
    box-shadow: 0 15px 40px rgba(31,65,114,.07);
}}

.panel-heading {{
    color: #15366c;
    font-size: 22px;
    font-weight: 900;
    text-align: center;
    margin-bottom: 5px;
}}

.panel-description {{
    color: #8290a7;
    font-size: 13px;
    text-align: center;
    margin-bottom: 20px;
}}

[data-testid="stFileUploader"] {{
    border: 2px dashed #bfdbfe;
    border-radius: 19px;
    background: #f7faff;
    padding: 14px;
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

/* Buttons */

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

/* Results */

.result-header {{
    background: linear-gradient(120deg,#102d61,#2563eb);
    border-radius: 23px;
    padding: 28px;
    color: white;
    margin-top: 30px;
    text-align: center;
    box-shadow: 0 12px 30px rgba(37,99,235,.14);
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
    margin: 8px 0 0;
}}

.score-card {{
    background: white;
    border: 1px solid #e5edf8;
    border-radius: 22px;
    padding: 25px 15px;
    text-align: center;
    box-shadow: 0 12px 30px rgba(31,65,114,.06);
    height: 100%;
}}

.score-circle {{
    width: 150px;
    height: 150px;
    border-radius: 50%;
    margin: 15px auto;
    display: flex;
    align-items: center;
    justify-content: center;
    position: relative;
}}

.score-circle:before {{
    content: "";
    position: absolute;
    width: 124px;
    height: 124px;
    background: white;
    border-radius: 50%;
}}

.score-value {{
    position: relative;
    z-index: 1;
    color: #1745a0;
    font-size: 37px;
    font-weight: 900;
}}

.score-label {{
    color: #74849d;
    font-size: 13px;
    font-weight: 700;
}}

.score-status {{
    display: inline-block;
    background: #eff6ff;
    color: #2563eb;
    border-radius: 30px;
    padding: 7px 17px;
    font-size: 12px;
    font-weight: 800;
}}

.result-card {{
    background: white;
    border: 1px solid #e7edf7;
    border-radius: 18px;
    padding: 21px;
    min-height: 160px;
    margin-bottom: 14px;
    box-shadow: 0 5px 20px rgba(31,65,114,.04);
    transition: all .2s ease;
}}

.result-card:hover {{
    border-color: #d5e5fb;
    box-shadow: 0 10px 25px rgba(37,99,235,.07);
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
    padding: 7px 0;
    border-bottom: 1px solid #f0f3f8;
    text-align: {'right' if ar else 'left'};
    overflow-wrap: anywhere;
}}

.result-item:last-child {{
    border-bottom: none;
}}

.footer {{
    border-top: 1px solid #e8eef7;
    padding-top: 22px;
    margin-top: 50px;
    text-align: center;
    color: #8795aa;
    font-size: 12px;
}}

.footer strong {{
    color: #2563eb;
    font-size: 15px;
}}

.analysis-note {{
    color: #8492a8;
    font-size: 11px;
    text-align: center;
    margin-top: 15px;
}}

/* Mobile */

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
        font-size: 24px;
    }}

    .feature-card {{
        padding: 17px 10px;
        min-height: 155px;
    }}

    .result-card {{
        padding: 16px;
    }}
}}
</style>
""", unsafe_allow_html=True)

# =========================
# LANGUAGE SWITCH BUTTON
# =========================

top_space, lang_col = st.columns([8, 1])

with lang_col:
    button_label = "EN" if ar else "عربي"

    if st.button(
        button_label,
        key="language_toggle",
        use_container_width=True
    ):
        st.session_state.lang = "en" if ar else "ar"
        st.rerun()

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

