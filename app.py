import os
import time
import io

import streamlit as st
import pypdf
import docx

from google import genai
from dotenv import load_dotenv

# PDF Generation Imports
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors


# ============================================================
# LOAD ENVIRONMENT VARIABLES & PAGE CONFIG
# ============================================================

load_dotenv()

st.set_page_config(
    page_title="WSA Educational Test Series & Paper Builder",
    page_icon="🎓",
    layout="wide"
)


# ============================================================
# CUSTOM CSS
# ============================================================

st.markdown(
    """
    <style>
        .stApp {
            background-color: #F8FAFC;
        }

        .hero-container {
            background: linear-gradient(
                135deg,
                #1E3A8A 0%,
                #3B82F6 100%
            );
            padding: 28px 32px;
            border-radius: 12px;
            color: white;
            box-shadow: 0 4px 12px rgba(30, 58, 138, 0.15);
            margin-bottom: 25px;
        }

        .hero-title {
            font-size: 2.3rem;
            font-weight: 800;
            margin: 0;
            letter-spacing: -0.5px;
        }

        .hero-subtitle {
            font-size: 1.05rem;
            opacity: 0.9;
            margin-top: 8px;
            margin-bottom: 0;
        }

        .edu-card {
            background-color: #FFFFFF;
            border: 1px solid #E2E8F0;
            border-radius: 10px;
            padding: 20px;
            box-shadow: 0 2px 6px rgba(0,0,0,0.03);
            margin-bottom: 20px;
        }

        .branding-card {
            background: linear-gradient(
                135deg,
                #0F172A 0%,
                #1E293B 100%
            );
            color: #F8FAFC;
            padding: 16px;
            border-radius: 10px;
            border-left: 4px solid #3B82F6;
            text-align: center;
            margin-top: 25px;
        }

        .branding-name {
            font-size: 1.05rem;
            font-weight: 700;
            color: #60A5FA;
            margin-bottom: 4px;
        }

        .branding-tag {
            font-size: 0.8rem;
            color: #94A3B8;
            letter-spacing: 0.5px;
        }

        .stButton > button {
            background: linear-gradient(
                135deg,
                #2563EB 0%,
                #1D4ED8 100%
            );
            color: white;
            font-weight: 600;
            border-radius: 8px;
            border: none;
            padding: 10px 24px;
        }
    </style>
    """,
    unsafe_allow_html=True
)


# ============================================================
# PDF GENERATOR FUNCTION (WITH APP & USER BRANDING)
# ============================================================

def create_pdf_from_text(title, content):
    """Converts markdown/text content into a downloadable PDF binary stream with app and author branding."""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        rightMargin=40,
        leftMargin=40,
        topMargin=40,
        bottomMargin=40
    )

    styles = getSampleStyleSheet()

    # Branding & Header Styles
    app_meta_style = ParagraphStyle(
        'AppMetaStyle',
        parent=styles['Normal'],
        fontSize=8,
        textColor=colors.HexColor('#64748B'),
        alignment=1,
        spaceAfter=10
    )

    title_style = ParagraphStyle(
        'TitleStyle',
        parent=styles['Heading1'],
        fontSize=18,
        textColor=colors.HexColor('#1E3A8A'),
        spaceAfter=15,
        alignment=1
    )

    body_style = ParagraphStyle(
        'BodyStyle',
        parent=styles['Normal'],
        fontSize=10,
        leading=14,
        textColor=colors.HexColor('#1E293B'),
        spaceAfter=8
    )

    footer_style = ParagraphStyle(
        'FooterStyle',
        parent=styles['Normal'],
        fontSize=9,
        textColor=colors.HexColor('#2563EB'),
        alignment=1,
        spaceBefore=15
    )

    # 1. Header Metadata & Title
    story = [
        Paragraph("<b>Generated via WSA Educational Test Series &amp; Paper Builder (AI-Powered)</b>", app_meta_style),
        Spacer(1, 4),
        Paragraph(title, title_style),
        Spacer(1, 10)
    ]

    # 2. Main Content Formatting
    lines = content.split('\n')
    for line in lines:
        clean_line = line.strip()
        if clean_line:
            clean_line = clean_line.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            clean_line = clean_line.replace('**', '<b>', 1)
            if '**' in clean_line:
                clean_line = clean_line.replace('**', '</b>', 1)

            story.append(Paragraph(clean_line, body_style))
        else:
            story.append(Spacer(1, 6))

    # 3. Footer Branding & Credit
    story.append(Spacer(1, 15))
    story.append(Paragraph("_________________________________________________________________________________", app_meta_style))
    story.append(Paragraph("<b>Designed by Waheed Ali Hamouzai</b> • WSA Educational Community", footer_style))

    doc.build(story)
    buffer.seek(0)
    return buffer


# ============================================================
# FILE EXTRACTION FUNCTION
# ============================================================

def process_uploaded_file(uploaded_file):
    if uploaded_file is None:
        return "", None

    filename = uploaded_file.name.lower()

    if filename.endswith(".txt"):
        try:
            return uploaded_file.getvalue().decode("utf-8").strip(), None
        except Exception:
            return "", None

    elif filename.endswith(".docx"):
        try:
            doc = docx.Document(uploaded_file)
            full_text = [p.text for p in doc.paragraphs if p.text]
            return "\n".join(full_text).strip(), None
        except Exception:
            return "", None

    elif filename.endswith(".pdf"):
        text = ""
        try:
            reader = pypdf.PdfReader(uploaded_file)
            for page in reader.pages:
                extracted = page.extract_text()
                if extracted:
                    text += extracted + "\n"
            text = text.strip()
        except Exception:
            text = ""

        if text:
            if len(text) > 15000:
                text = text[:15000]
            return text, None
        else:
            uploaded_file.seek(0)
            return "", {
                "mime_type": "application/pdf",
                "data": uploaded_file.getvalue()
            }

    elif filename.endswith((".png", ".jpg", ".jpeg")):
        mime_type = "image/png" if filename.endswith(".png") else "image/jpeg"
        uploaded_file.seek(0)
        return "", {
            "mime_type": mime_type,
            "data": uploaded_file.getvalue()
        }

    return "", None


# ============================================================
# GEMINI API CALL WITH FALLBACK
# ============================================================

def call_gemini_with_retry(
    api_key,
    prompt_text,
    file_attachments=None,
    temperature=0.3
):
    client = genai.Client(api_key=api_key)

    candidate_models = [
        "gemini-2.5-flash",
        "gemini-2.0-flash",
        "gemini-1.5-flash"
    ]

    last_error = ""

    for model_name in candidate_models:
        try:
            contents = []
            if file_attachments:
                for attachment in file_attachments:
                    if attachment:
                        contents.append(attachment)

            contents.append(prompt_text)

            response = client.models.generate_content(
                model=model_name,
                contents=contents,
                config={"temperature": temperature}
            )

            if response and response.text:
                return response.text

        except Exception as e:
            last_error = str(e)
            if "404" in last_error or "NOT_FOUND" in last_error or "not found" in last_error.lower():
                continue
            time.sleep(1)

    raise Exception(
        f"Model connection failed. Check API Key.\n\nError: {last_error}"
    )


# ============================================================
# AI GENERATION & EVALUATION LOGIC
# ============================================================

def generate_test_paper(
    api_key, topic, uploaded_pdf, test_type,
    mcq_count, short_count, long_count, diff_level
):
    pdf_text, pdf_part = process_uploaded_file(uploaded_pdf)
    file_attachments = [pdf_part] if pdf_part else None

    prompt = f"""
You are an expert examiner for educational boards in Pakistan.

Create a professional examination paper.

TARGET TEST CATEGORY: {test_type}
TOPIC / SUBJECT: {topic}
DIFFICULTY LEVEL: {diff_level}

QUESTION COUNTS:
MCQs: {mcq_count}
SHORT QUESTIONS: {short_count}
LONG QUESTIONS: {long_count}

REQUIREMENTS:
1. MCQs must have 4 options (A, B, C, D).
2. Short and Long questions should be clear.
3. Provide a complete ANSWER KEY at the bottom.

{f"REFERENCE TEXT:\n{pdf_text}" if pdf_text else ""}
"""

    return call_gemini_with_retry(
        api_key=api_key,
        prompt_text=prompt,
        file_attachments=file_attachments,
        temperature=0.3
    )


def evaluate_student_answers(
    api_key, paper_text, paper_attachment,
    answers_text, answers_attachment
):
    file_attachments = []
    if paper_attachment:
        file_attachments.append(paper_attachment)
    if answers_attachment:
        file_attachments.append(answers_attachment)

    prompt = f"""
You are an experienced examiner.

Evaluate the student's answer sheet against the provided question paper and answer key.

QUESTION PAPER / ANSWER KEY:
{paper_text if paper_text else "Attached as file/image."}

STUDENT ANSWERS:
{answers_text if answers_text else "Attached as file/image."}

Provide a structured evaluation report:
- Total Marks & Obtained Marks
- Percentage & Grade
- Question-wise Analysis
- Strengths & Weaknesses
"""

    return call_gemini_with_retry(
        api_key=api_key,
        prompt_text=prompt,
        file_attachments=file_attachments if file_attachments else None,
        temperature=0.2
    )


# ============================================================
# SIDEBAR & HEADER
# ============================================================

with st.sidebar:
    st.header("⚙️ Settings")
    env_api_key = os.getenv("GEMINI_API_KEY", "")
    api_key_input = st.text_input(
        "Enter Gemini API Key",
        value=env_api_key,
        type="password"
    )
    api_key = api_key_input or env_api_key

st.markdown(
    """<div class="hero-container">
<div class="hero-title">🎓 WSA Educational Test Series & Paper Builder</div>
<div class="hero-subtitle">AI-powered exam paper generation and answer evaluation.</div>
</div>""",
    unsafe_allow_html=True
)

tab1, tab2 = st.tabs(["📝 Generate Test Paper", "📊 Evaluate Student Answers"])


# ============================================================
# TAB 1 — GENERATE PAPER
# ============================================================

with tab1:
    st.markdown('<div class="edu-card">', unsafe_allow_html=True)
    st.subheader("Create New Test Paper")

    col1, col2 = st.columns(2)

    with col1:
        test_type = st.selectbox(
            "Exam Category",
            [
                "General Competitive Test (CTSP / SBK / NTS)",
                "BPSC",
                "School / College (9th–12th)",
                "CS / IT Screening Test",
                "Custom Mock Test"
            ]
        )
        topic = st.text_input("Topic / Subject", placeholder="e.g., Computer Science")
        difficulty = st.slider("Difficulty Level", 1, 5, 3)
        diff_labels = {1: "Very Easy", 2: "Easy", 3: "Medium", 4: "Hard", 5: "Very Hard"}
        diff_level = diff_labels[difficulty]

    with col2:
        uploaded_pdf = st.file_uploader("Upload Reference Document (Optional)", type=["pdf", "docx", "txt"])
        mcq_count = st.number_input("Number of MCQs", 0, 50, 10)
        short_count = st.number_input("Number of Short Questions", 0, 20, 5)
        long_count = st.number_input("Number of Long Questions", 0, 10, 2)

    st.markdown("</div>", unsafe_allow_html=True)

    if st.button("🚀 Generate Test Paper", use_container_width=True):
        if not api_key:
            st.error("⚠️ Please enter API key.")
        elif not topic.strip():
            st.error("⚠️ Please enter topic.")
        else:
            with st.spinner("Generating test paper..."):
                try:
                    res = generate_test_paper(
                        api_key, topic, uploaded_pdf, test_type,
                        mcq_count, short_count, long_count, diff_level
                    )
                    st.session_state["generated_paper"] = res
                except Exception as e:
                    st.error(f"❌ Error: {e}")

    if "generated_paper" in st.session_state:
        st.divider()
        st.subheader("📄 Generated Test Paper")
        st.markdown(st.session_state["generated_paper"])

        c1, c2 = st.columns(2)
        with c1:
            st.download_button(
                "⬇️ Download Text (.txt)",
                data=st.session_state["generated_paper"],
                file_name="WSA_Test_Paper.txt",
                mime="text/plain",
                use_container_width=True
            )
        with c2:
            pdf_bytes = create_pdf_from_text("WSA Examination Paper", st.session_state["generated_paper"])
            st.download_button(
                "⬇️ Download PDF (.pdf)",
                data=pdf_bytes,
                file_name="WSA_Test_Paper.pdf",
                mime="application/pdf",
                use_container_width=True
            )


# ============================================================
# TAB 2 — EVALUATE ANSWERS (TEXT + PDF DOWNLOAD)
# ============================================================

with tab2:
    st.subheader("📊 Evaluate Student Answers")
    default_paper = st.session_state.get("generated_paper", "")

    col_paper, col_answer = st.columns(2)

    with col_paper:
        st.markdown("### 1️⃣ Question Paper / Answer Key")
        paper_file = st.file_uploader(
            "Upload Paper (PDF, DOCX, TXT, Image)",
            type=["pdf", "docx", "txt", "png", "jpg", "jpeg"],
            key="p_up"
        )
        question_paper_text = st.text_area(
            "Or Paste Text",
            value=default_paper,
            height=200
        )

    with col_answer:
        st.markdown("### 2️⃣ Student Answer Sheet")
        student_file = st.file_uploader(
            "Upload Answers (PDF, DOCX, TXT, Image)",
            type=["pdf", "docx", "txt", "png", "jpg", "jpeg"],
            key="s_up"
        )
        student_answers_text = st.text_area(
            "Or Paste Text",
            height=200
        )

    if st.button("📊 Evaluate Answers", use_container_width=True):
        if not api_key:
            st.error("⚠️ Please enter API key.")
        else:
            p_text, p_part = process_uploaded_file(paper_file)
            final_p_text = p_text or question_paper_text.strip()

            a_text, a_part = process_uploaded_file(student_file)
            final_a_text = a_text or student_answers_text.strip()

            if not (final_p_text or p_part):
                st.error("⚠️ Question paper is missing.")
            elif not (final_a_text or a_part):
                st.error("⚠️ Student answers are missing.")
            else:
                with st.spinner("Evaluating student answers..."):
                    try:
                        eval_res = evaluate_student_answers(
                            api_key, final_p_text, p_part, final_a_text, a_part
                        )
                        st.session_state["evaluation"] = eval_res
                    except Exception as e:
                        st.error(f"❌ Error: {e}")

    if "evaluation" in st.session_state:
        st.divider()
        st.subheader("📋 Evaluation Result")
        st.markdown(st.session_state["evaluation"])

        col_d1, col_d2 = st.columns(2)
        with col_d1:
            st.download_button(
                "⬇️ Download Text Report (.txt)",
                data=st.session_state["evaluation"],
                file_name="WSA_Student_Evaluation.txt",
                mime="text/plain",
                use_container_width=True
            )
        with col_d2:
            eval_pdf = create_pdf_from_text("Student Evaluation Report", st.session_state["evaluation"])
            st.download_button(
                "⬇️ Download PDF Report (.pdf)",
                data=eval_pdf,
                file_name="WSA_Student_Evaluation.pdf",
                mime="application/pdf",
                use_container_width=True
            )


# ============================================================
# FOOTER
# ============================================================

st.markdown(
    """<div class="branding-card">
<div class="branding-name">Designed by Waheed Ali Hamouzai</div>
<div class="branding-tag">WSA Educational Community • AI Powered Learning</div>
</div>""",
    unsafe_allow_html=True
)
