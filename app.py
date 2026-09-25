import os
import time
import io

import streamlit as st
import pypdf
import docx

from openai import OpenAI
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
# PDF GENERATOR FUNCTION
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

    story = [
        Paragraph("<b>Generated via WSA Educational Test Series &amp; Paper Builder (AI-Powered)</b>", app_meta_style),
        Spacer(1, 4),
        Paragraph(title, title_style),
        Spacer(1, 10)
    ]

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
        return ""

    filename = uploaded_file.name.lower()

    if filename.endswith(".txt"):
        try:
            return uploaded_file.getvalue().decode("utf-8").strip()
        except Exception:
            return ""

    elif filename.endswith(".docx"):
        try:
            doc = docx.Document(uploaded_file)
            full_text = [p.text for p in doc.paragraphs if p.text]
            return "\n".join(full_text).strip()
        except Exception:
            return ""

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

        if len(text) > 15000:
            text = text[:15000]
        return text

    return ""


# ============================================================
# OPENAI API CALL FUNCTION
# ============================================================

def call_openai_api(api_key, prompt_text, system_instruction="You are an expert educational examiner.", temperature=0.3):
    client = OpenAI(api_key=api_key)
    
    response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[
            {"role": "system", "content": system_instruction},
            {"role": "user", "content": prompt_text}
        ],
        temperature=temperature
    )
    
    return response.choices[0].message.content


# ============================================================
# AI GENERATION & EVALUATION LOGIC
# ============================================================

def generate_test_paper(
    api_key, topic, uploaded_pdf, test_type, language,
    mcq_count, short_count, long_count, diff_level
):
    pdf_text = process_uploaded_file(uploaded_pdf)

    language_instruction = ""
    if language == "Urdu":
        language_instruction = "Generate the complete test paper strictly in URDU language."
    elif language == "Bilingual (English + Urdu)":
        language_instruction = (
            "Generate the test paper in BILINGUAL format (English followed by Urdu translation for each question and option)."
        )
    else:
        language_instruction = "Generate the test paper in ENGLISH language."

    prompt = f"""
Create a professional examination paper.

TARGET TEST CATEGORY: {test_type}
TOPIC / SUBJECT: {topic}
DIFFICULTY LEVEL: {diff_level}
LANGUAGE MODE: {language}

LANGUAGE REQUIREMENT:
{language_instruction}

QUESTION COUNTS:
MCQs: {mcq_count}
SHORT QUESTIONS: {short_count}
LONG QUESTIONS: {long_count}

REQUIREMENTS:
1. MCQs must have 4 options (A, B, C, D).
2. Short and Long questions should be clear and well-structured.
3. Provide a complete ANSWER KEY at the bottom.

{f"REFERENCE TEXT:\n{pdf_text}" if pdf_text else ""}
"""

    return call_openai_api(
        api_key=api_key,
        prompt_text=prompt,
        system_instruction="You are an expert examiner for educational boards and competitive testing services capable of generating test papers in English, Urdu, and Bilingual formats.",
        temperature=0.3
    )


def evaluate_student_answers(
    api_key, paper_text, answers_text
):
    prompt = f"""
Evaluate the student's answer sheet against the provided question paper and answer key.

QUESTION PAPER / ANSWER KEY:
{paper_text}

STUDENT ANSWERS:
{answers_text}

Provide a structured evaluation report:
- Total Marks & Obtained Marks
- Percentage & Grade
- Question-wise Analysis
- Strengths & Weaknesses
"""

    return call_openai_api(
        api_key=api_key,
        prompt_text=prompt,
        system_instruction="You are an experienced examiner evaluating student answers accurately and providing constructive feedback.",
        temperature=0.2
    )


def analyze_random_test(api_key, test_content, additional_context=""):
    prompt = f"""
Analyze the following random test paper, solved sheet, or test result provided by the user.

USER CONTEXT / GOAL:
{additional_context if additional_context else "General test analysis and performance diagnostic."}

TEST DATA / CONTENT:
{test_content}

Provide a comprehensive Diagnostic & Improvement Report structured as follows:
1. 📌 **Executive Performance Overview**: High-level estimation of score, accuracy, or completion quality.
2. 🎯 **Key Weaknesses & Knowledge Gaps**: Identify exact topics, question types, or concepts where performance is lacking.
3. 🔍 **Priority Focus Areas**: Highlight top 3 critical subjects/topics the student MUST prioritize immediately.
4. 🚀 **Actionable Improvement Strategy**: Step-by-step study recommendations, practice methods, and revision plan.
5. 💡 **Recommended Resources & Next Steps**: Suggested topics to solve next or key formulas/concepts to memorize.
"""

    return call_openai_api(
        api_key=api_key,
        prompt_text=prompt,
        system_instruction="You are a senior academic mentor and diagnostic expert specializing in test analysis and student performance optimization.",
        temperature=0.3
    )


# ============================================================
# SIDEBAR & HEADER
# ============================================================

with st.sidebar:
    st.header("⚙️ Settings")
    env_api_key = os.getenv("OPENAI_API_KEY", "")
    api_key_input = st.text_input(
        "Enter OpenAI API Key (sk-...)",
        value=env_api_key,
        type="password"
    )
    api_key = api_key_input or env_api_key

st.markdown(
    """<div class="hero-container">
<div class="hero-title">🎓 WSA Educational Test Series & Paper Builder</div>
<div class="hero-subtitle">AI-powered exam paper generation, student evaluation, and performance diagnostics.</div>
</div>""",
    unsafe_allow_html=True
)

tab1, tab2, tab3 = st.tabs([
    "📝 Generate Test Paper", 
    "📊 Evaluate Student Answers", 
    "🎯 Quick Test Diagnostic & Focus Areas"
])


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
        topic = st.text_input("Topic / Subject", placeholder="e.g., Computer Science / Pak Studies")
        language = st.selectbox(
            "Language Mode",
            ["English", "Urdu", "Bilingual (English + Urdu)"]
        )
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
            st.error("⚠️ Please enter a valid OpenAI API key in the sidebar.")
        elif not topic.strip():
            st.error("⚠️ Please enter a topic or subject name.")
        else:
            with st.spinner("Generating test paper via AI... Please wait."):
                try:
                    res = generate_test_paper(
                        api_key, topic, uploaded_pdf, test_type, language,
                        mcq_count, short_count, long_count, diff_level
                    )
                    st.session_state["generated_paper"] = res
                except Exception as e:
                    st.error(f"❌ Error occurred: {e}")

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
# TAB 2 — EVALUATE ANSWERS
# ============================================================

with tab2:
    st.subheader("📊 Evaluate Student Answers")
    default_paper = st.session_state.get("generated_paper", "")

    col_paper, col_answer = st.columns(2)

    with col_paper:
        st.markdown("### 1️⃣ Question Paper / Answer Key")
        paper_file = st.file_uploader(
            "Upload Paper (PDF, DOCX, TXT)",
            type=["pdf", "docx", "txt"],
            key="p_up"
        )
        question_paper_text = st.text_area(
            "Or Paste Text Directly",
            value=default_paper,
            height=200,
            key="q_paper_text"
        )

    with col_answer:
        st.markdown("### 2️⃣ Student Answer Sheet")
        student_file = st.file_uploader(
            "Upload Answers (PDF, DOCX, TXT)",
            type=["pdf", "docx", "txt"],
            key="s_up"
        )
        student_answers_text = st.text_area(
            "Or Paste Text Directly",
            height=200,
            key="s_answers_text"
        )

    if st.button("📊 Evaluate Answers", use_container_width=True):
        if not api_key:
            st.error("⚠️ Please enter a valid OpenAI API key in the sidebar.")
        else:
            p_text = process_uploaded_file(paper_file)
            final_p_text = p_text or question_paper_text.strip()

            a_text = process_uploaded_file(student_file)
            final_a_text = a_text or student_answers_text.strip()

            if not final_p_text:
                st.error("⚠️ Question paper content is missing. Please upload or paste text.")
            elif not final_a_text:
                st.error("⚠️ Student answer content is missing. Please upload or paste text.")
            else:
                with st.spinner("Evaluating student answers via AI... Please wait."):
                    try:
                        eval_res = evaluate_student_answers(
                            api_key, final_p_text, final_a_text
                        )
                        st.session_state["evaluation"] = eval_res
                    except Exception as e:
                        st.error(f"❌ Error occurred: {e}")

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
# TAB 3 — RANDOM TEST DIAGNOSTIC & IMPROVEMENT
# ============================================================

with tab3:
    st.subheader("🎯 Test Performance Diagnostic & Focus Area Planner")
    st.markdown("Upload or paste any random solved test, question paper, or result card to receive a detailed breakdown of strengths, weaknesses, focus areas, and improvement plans.")

    col_diag1, col_diag2 = st.columns(2)

    with col_diag1:
        st.markdown("### 📄 Test File / Raw Data")
        random_test_file = st.file_uploader(
            "Upload Test or Result Document (PDF, DOCX, TXT)",
            type=["pdf", "docx", "txt"],
            key="random_test_up"
        )
        random_test_text = st.text_area(
            "Or Paste Raw Test Content / Scores Directly",
            height=200,
            key="random_test_text",
            placeholder="e.g., Question 1: Incorrect answer selected...\nQuestion 2: Right...\nOR paste raw quiz content here."
        )

    with col_diag2:
        st.markdown("### ⚙️ Context & Target Goals (Optional)")
        user_context = st.text_area(
            "Target Exam / Student Goal",
            height=270,
            key="user_context",
            placeholder="e.g., Preparing for BPSC Computer Science / SBK Screening Test. Target score is 80%+."
        )

    if st.button("🔍 Analyze Test & Generate Focus Plan", use_container_width=True):
        if not api_key:
            st.error("⚠️ Please enter a valid OpenAI API key in the sidebar.")
        else:
            extracted_test = process_uploaded_file(random_test_file)
            final_test_content = extracted_test or random_test_text.strip()

            if not final_test_content:
                st.error("⚠️ Please upload a test file or paste text to analyze.")
            else:
                with st.spinner("Analyzing test data and designing personalized improvement plan..."):
                    try:
                        diag_res = analyze_random_test(
                            api_key, final_test_content, user_context
                        )
                        st.session_state["diagnostic_analysis"] = diag_res
                    except Exception as e:
                        st.error(f"❌ Error occurred: {e}")

    if "diagnostic_analysis" in st.session_state:
        st.divider()
        st.subheader("📈 Diagnostic Report & Improvement Plan")
        st.markdown(st.session_state["diagnostic_analysis"])

        col_rd1, col_rd2 = st.columns(2)
        with col_rd1:
            st.download_button(
                "⬇️ Download Text Diagnostic (.txt)",
                data=st.session_state["diagnostic_analysis"],
                file_name="WSA_Test_Diagnostic.txt",
                mime="text/plain",
                use_container_width=True
            )
        with col_rd2:
            diag_pdf = create_pdf_from_text("Test Diagnostic & Improvement Plan", st.session_state["diagnostic_analysis"])
            st.download_button(
                "⬇️ Download PDF Diagnostic (.pdf)",
                data=diag_pdf,
                file_name="WSA_Test_Diagnostic.pdf",
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
