import os
import time
import streamlit as st
import pypdf
from google import genai
from google.genai import types
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Page configuration
st.set_page_config(
    page_title="WSA Educational Test Series & Paper Builder",
    page_icon="🎓",
    layout="wide"
)

# --- MODERN EDUCATIONAL THEME (CUSTOM CSS) ---
st.markdown("""
<style>
    .stApp { background-color: #F8FAFC; }
    
    .hero-container {
        background: linear-gradient(135deg, #1E3A8A 0%, #3B82F6 100%);
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
        background: linear-gradient(135deg, #0F172A 0%, #1E293B 100%);
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
    
    .stButton>button {
        background: linear-gradient(135deg, #2563EB 0%, #1D4ED8 100%);
        color: white;
        font-weight: 600;
        border-radius: 8px;
        border: none;
        padding: 10px 24px;
        box-shadow: 0 2px 4px rgba(37, 99, 235, 0.2);
        transition: all 0.2s ease;
    }
    .stButton>button:hover {
        background: linear-gradient(135deg, #1D4ED8 0%, #1E40AF 100%);
        box-shadow: 0 4px 8px rgba(37, 99, 235, 0.3);
    }
</style>
""", unsafe_allow_html=True)


def extract_pdf_text(uploaded_file):
    """PDF text extraction helper."""
    try:
        reader = pypdf.PdfReader(uploaded_file)
        text = ""
        for page in reader.pages:
            extracted = page.extract_text()
            if extracted:
                text += extracted + "\n"
        return text.strip()
    except Exception:
        return ""


def call_gemini_with_retry(client, contents, temperature=0.3, max_tokens=3500, retries=3):
    """503 high demand error se bachne ke liye auto-retry mechanism."""
    for attempt in range(retries):
        try:
            response = client.models.generate_content(
                model='gemini-1.5-flash',  # Most stable production model
                contents=contents,
                config=types.GenerateContentConfig(
                    temperature=temperature,
                    max_output_tokens=max_tokens,
                )
            )
            return response.text
        except Exception as e:
            err_str = str(e)
            if "503" in err_str or "UNAVAILABLE" in err_str or "429" in err_str:
                if attempt < retries - 1:
                    time.sleep(2 * (attempt + 1))  # Pause for 2s then 4s before retrying
                    continue
            raise e


def generate_test_paper(api_key, topic, uploaded_pdf, test_type, mcq_count, short_count, long_count, diff_level):
    """Gemini 1.5 Flash ke zariye complete Question Paper aur Answer Key generate karna."""
    try:
        client = genai.Client(api_key=api_key)

        prompt = f"""
You are an expert examiner for competitive testing agencies in Balochistan/Pakistan (such as CTSP, SBK, BPSC, NTS) and secondary/higher education boards.

Generate a highly structured, professional Question Paper along with a complete Answer Key based on the provided details.

Target Test Category/Role: {test_type}
Topic / Subject Title: {topic}
Difficulty Level: {diff_level}
Paper Pattern Requirements:
- MCQs: {mcq_count} questions (with 4 options: A, B, C, D)
- Short Questions: {short_count} questions
- Long/Descriptive Questions: {long_count} questions

Use standard subject knowledge matching the topic, test level, and any attached reference document.

Structure your output into TWO clearly separated main sections using Markdown formatting:

# SECTION A: QUESTION PAPER
(Format this professionally as a printable exam paper with headers like Marks, Subject, Time Allowed, and clear numbering for questions.)

---

# SECTION B: ANSWER KEY & MARKING SCHEME
(Provide accurate correct choices for all MCQs, concise point-by-point model answers for Short Questions, and main key evaluation points for Long Questions.)
"""

        contents = []
        if uploaded_pdf is not None:
            bytes_data = uploaded_pdf.getvalue()
            contents.append(types.Part.from_bytes(data=bytes_data, mime_type="application/pdf"))
            
        contents.append(prompt)

        return call_gemini_with_retry(client, contents, temperature=0.3, max_tokens=3500)
    except Exception as e:
        return f"Error generating test paper: {str(e)}"


def evaluate_student_answers(api_key, question_paper, student_answers):
    """Student ke answers ko grade karke detailed marks aur feedback dena."""
    try:
        client = genai.Client(api_key=api_key)

        prompt = f"""
You are an experienced academic evaluator. Grade the following student's submission against the provided Question Paper.

Question Paper / Test Reference:
\"\"\"
{question_paper}
\"\"\"

Student Submitted Answers:
\"\"\"
{student_answers}
\"\"\"

Please provide a structured grading report in clean Markdown:

1. **Overall Performance & Score Summary**: Total estimated marks obtained vs maximum marks.
2. **Detailed Breakdown**: Question-by-question evaluation indicating correct answers, partial credit, and mistakes.
3. **Key Weak Areas**: List topics or concepts where the student needs improvement.
4. **Actionable Suggestions**: 2-3 specific recommendations for better test preparation.
"""

        contents = [prompt]
        return call_gemini_with_retry(client, contents, temperature=0.2, max_tokens=2500)
    except Exception as e:
        return f"Error evaluating submission: {str(e)}"


# --- HERO HEADER ---
st.markdown("""
<div class="hero-container">
    <div class="hero-title">🎓 WSA Local Test Series & Paper Builder</div>
    <div class="hero-subtitle">Smart AI Exam Suite — Generate Custom Mock Tests, Answer Keys & Automated Student Evaluations</div>
</div>
""", unsafe_allow_html=True)

# --- SIDEBAR CONFIGURATION & ABOUT BRANDING ---
with st.sidebar:
    st.header("⚙️ System Setup")
    env_api_key = os.getenv("GEMINI_API_KEY", "")
    user_api_key = st.text_input("Gemini API Key", value=env_api_key, type="password", help="Enter your Gemini API Key or configure it in .env")
    
    st.divider()
    st.markdown("### 🎯 Exam Category")
    test_type = st.selectbox(
        "Select Target Testing Body / Level:",
        [
            "General Competitive Test (CTSP / SBK / NTS)",
            "BPSC Screening & General Knowledge",
            "School / College Board Exam (9th-12th Class)",
            "Computer Science & IT Skill Screening",
            "Custom Mock Screening Series"
        ]
    )
    
    st.divider()
    
    # --- ABOUT & DESIGNER BRANDING ---
    st.markdown("""
    <div class="branding-card">
        <div style="font-size: 0.75rem; text-transform: uppercase; color: #94A3B8; letter-spacing: 1px;">Platform Lead</div>
        <div class="branding-name">Designed by Waheed Ali Hamouzai</div>
        <div class="branding-tag">WSA Educational Community</div>
    </div>
    """, unsafe_allow_html=True)

api_key = user_api_key or os.getenv("GEMINI_API_KEY")

# --- NAVIGATION TABS ---
tab1, tab2 = st.tabs(["📄 1. Test Paper & Answer Key Generator", "📊 2. Student Answer Sheet Evaluator"])

# --- TAB 1: PAPER BUILDER ---
with tab1:
    st.markdown("##### 📝 Configure Exam Parameters")
    
    col1, col2 = st.columns([1, 1])
    
    with col1:
        topic = st.text_input("Subject / Topic Title:", placeholder="e.g. C++ Programming Fundamentals, Everyday Science, General Knowledge")
        diff_level = st.select_slider("Select Difficulty Level:", options=["Easy", "Medium", "Hard", "Advanced Competitive"])
        
        uploaded_pdf = st.file_uploader("Upload Syllabus / Chapter PDF (Optional):", type=["pdf"])
        if uploaded_pdf:
            extracted_text = extract_pdf_text(uploaded_pdf)
            if len(extracted_text) > 0:
                st.success(f"Text PDF detected ({len(extracted_text)} characters extracted).")
            else:
                st.info("Scanned/Image PDF detected. Sending document directly to Gemini Flash for OCR processing.")

    with col2:
        st.markdown("**Question Distribution:**")
        mcq_count = st.number_input("Number of MCQs:", min_value=0, max_value=50, value=10)
        short_count = st.number_input("Number of Short Questions:", min_value=0, max_value=20, value=3)
        long_count = st.number_input("Number of Long / Descriptive Questions:", min_value=0, max_value=10, value=1)

    st.divider()

    if st.button("🚀 Generate Exam Paper & Answer Key"):
        if not api_key:
            st.error("Please provide a Gemini API Key in the sidebar or setup your `.env` file.")
        elif not topic.strip():
            st.warning("Please enter a Subject / Topic Title.")
        else:
            with st.spinner("Generating Question Paper and Answer Key via Gemini 1.5 Flash..."):
                generated_result = generate_test_paper(
                    api_key=api_key,
                    topic=topic,
                    uploaded_pdf=uploaded_pdf,
                    test_type=test_type,
                    mcq_count=mcq_count,
                    short_count=short_count,
                    long_count=long_count,
                    diff_level=diff_level
                )
                
                st.session_state["last_generated_paper"] = generated_result
                
                st.markdown("### 📋 Exam Output")
                st.markdown(f"<div class='edu-card'>{generated_result}</div>", unsafe_allow_html=True)
                
                # Download Action
                st.download_button(
                    label="📥 Download Paper & Key (.txt)",
                    data=generated_result,
                    file_name=f"{topic.replace(' ', '_')}_Test_Paper.txt",
                    mime="text/plain"
                )

# --- TAB 2: ANSWER EVALUATOR ---
with tab2:
    st.markdown("##### 🔍 Evaluate Student Submissions")
    st.write("Paste the Question Paper / Marking Key along with the Student's Answers to get an automated grading report.")
    
    eval_col1, eval_col2 = st.columns(2)
    
    with eval_col1:
        default_paper = st.session_state.get("last_generated_paper", "")
        paper_text = st.text_area(
            "Reference Question Paper & Key:",
            value=default_paper,
            height=280,
            placeholder="Paste the original question paper or key here..."
        )
        
    with eval_col2:
        student_answers = st.text_area(
            "Student Submitted Responses:",
            height=280,
            placeholder="Paste the student's answers here (e.g., 1. A, 2. C, Short Ans 1: ...)..."
        )
        
    if st.button("📊 Evaluate Answer Sheet"):
        if not api_key:
            st.error("Please provide a Gemini API Key in the sidebar.")
        elif not paper_text.strip():
            st.warning("Please provide the reference Question Paper / Key.")
        elif not student_answers.strip():
            st.warning("Please paste the Student's Submitted Answers.")
        else:
            with st.spinner("Analyzing student responses and generating feedback report..."):
                evaluation_result = evaluate_student_answers(
                    api_key=api_key,
                    question_paper=paper_text,
                    student_answers=student_answers
                )
                
                st.markdown("### 🎯 Automated Grading Report")
                st.markdown(f"<div class='edu-card'>{evaluation_result}</div>", unsafe_allow_html=True)
