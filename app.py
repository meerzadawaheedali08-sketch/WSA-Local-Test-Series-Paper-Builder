import os
import streamlit as st
import PyPDF2 as pdf
from google import genai
from google.genai import types
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Page configuration
st.set_page_config(
    page_title="WSA Local Test Series & Paper Builder",
    page_icon="📝",
    layout="wide"
)

# Custom Styling
st.markdown("""
<style>
    .main-title { font-size: 2.2rem; font-weight: 700; color: #1E3A8A; margin-bottom: 0px; }
    .sub-title { font-size: 1.05rem; color: #4B5563; margin-bottom: 25px; }
    .stButton>button { background-color: #1E40AF; color: white; font-weight: bold; border-radius: 6px; }
    .paper-box { background-color: #F8FAFC; border: 1px solid #E2E8F0; padding: 20px; border-radius: 8px; }
</style>
""", unsafe_allow_html=True)


def extract_pdf_text(uploaded_file):
    """PDF file se text extract karne ke liye helper function."""
    try:
        reader = pdf.PdfReader(uploaded_file)
        text = ""
        for page in reader.pages:
            extracted = page.extract_text()
            if extracted:
                text += extracted + "\n"
        return text.strip()
    except Exception as e:
        st.error(f"Error reading PDF file: {str(e)}")
        return ""


def generate_test_paper(api_key, topic, syllabus_text, test_type, mcq_count, short_count, long_count, diff_level):
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

Reference Syllabus / Notes Text:
\"\"\"
{syllabus_text if syllabus_text else "Use standard subject knowledge matching the topic and test level."}
\"\"\"

Structure your output into TWO clearly separated main sections using Markdown formatting:

# SECTION A: QUESTION PAPER
(Format this professionally as a printable exam paper with headers like Marks, Subject, Time Allowed, and clear numbering for questions.)

---

# SECTION B: ANSWER KEY & MARKING SCHEME
(Provide accurate correct choices for all MCQs, concise point-by-point model answers for Short Questions, and main key evaluation points for Long Questions.)
"""

        response = client.models.generate_content(
            model='gemini-1.5-flash',
            contents=prompt,
            config=types.GenerateContentConfig(
                temperature=0.3,
                max_output_tokens=3000,
            )
        )
        return response.text
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

        response = client.models.generate_content(
            model='gemini-1.5-flash',
            contents=prompt,
            config=types.GenerateContentConfig(
                temperature=0.2,
                max_output_tokens=2500,
            )
        )
        return response.text
    except Exception as e:
        return f"Error evaluating submission: {str(e)}"


# --- MAIN UI ---
st.markdown("<p class='main-title'>📝 WSA Local Test Series & Paper Builder</p>", unsafe_allow_html=True)
st.markdown("<p class='sub-title'>Generate custom mock tests, answer keys, and evaluate student responses with Gemini 1.5 Flash.</p>", unsafe_allow_html=True)

# Sidebar Configuration
with st.sidebar:
    st.header("⚙️ API Configuration")
    env_api_key = os.getenv("GEMINI_API_KEY", "")
    user_api_key = st.text_input("Gemini API Key", value=env_api_key, type="password", help="Enter your Gemini API Key or set GEMINI_API_KEY in .env file")
    
    st.divider()
    st.markdown("### 🎯 Exam Type")
    test_type = st.selectbox(
        "Select Target Exam / Level:",
        [
            "General Competitive Test (CTSP / SBK / NTS)",
            "BPSC Screening & General Knowledge",
            "School / College Board Exam (9th-12th Class)",
            "Computer Science & IT Skill Screening",
            "Custom Mock Screening Series"
        ]
    )

api_key = user_api_key or os.getenv("GEMINI_API_KEY")

# Navigation Tabs
tab1, tab2 = st.tabs(["📄 1. Paper Builder", "📊 2. Student Answer Evaluator"])

# --- TAB 1: PAPER BUILDER ---
with tab1:
    st.subheader("Test Paper Parameters")
    
    col1, col2 = st.columns([1, 1])
    
    with col1:
        topic = st.text_input("Subject / Topic Title:", placeholder="e.g. C++ Programming Fundamentals, Everyday Science, General Knowledge")
        diff_level = st.select_slider("Difficulty Level:", options=["Easy", "Medium", "Hard", "Advanced Competitive"])
        
        uploaded_pdf = st.file_uploader("Upload Syllabus / Chapter PDF (Optional):", type=["pdf"])
        syllabus_text = ""
        if uploaded_pdf:
            syllabus_text = extract_pdf_text(uploaded_pdf)
            st.success(f"Extracted {len(syllabus_text)} characters from PDF.")

    with col2:
        st.markdown("**Question Distribution:**")
        mcq_count = st.number_input("Number of MCQs:", min_value=0, max_value=50, value=10)
        short_count = st.number_input("Number of Short Questions:", min_value=0, max_value=20, value=3)
        long_count = st.number_input("Number of Long/Descriptive Questions:", min_value=0, max_value=10, value=1)

    st.divider()

    if st.button("🚀 Generate Question Paper & Answer Key"):
        if not api_key:
            st.error("Please provide a Gemini API Key in the sidebar or setup your `.env` file.")
        elif not topic.strip():
            st.warning("Please enter a Subject / Topic Title.")
        else:
            with st.spinner("Generating Question Paper and Answer Key with Gemini Flash..."):
                generated_result = generate_test_paper(
                    api_key=api_key,
                    topic=topic,
                    syllabus_text=syllabus_text,
                    test_type=test_type,
                    mcq_count=mcq_count,
                    short_count=short_count,
                    long_count=long_count,
                    diff_level=diff_level
                )
                
                st.session_state["last_generated_paper"] = generated_result
                st.markdown("### 📋 Generated Paper & Key")
                st.markdown(generated_result)
                
                # Download Button
                st.download_button(
                    label="📥 Download Paper & Answer Key (.txt)",
                    data=generated_result,
                    file_name=f"{topic.replace(' ', '_')}_Test_Paper.txt",
                    mime="text/plain"
                )

# --- TAB 2: ANSWER EVALUATOR ---
with tab2:
    st.subheader("Evaluate Student Submission")
    st.write("Paste the Question Paper and the Student's Answers below to generate automated grading and feedback.")
    
    eval_col1, eval_col2 = st.columns(2)
    
    with eval_col1:
        default_paper = st.session_state.get("last_generated_paper", "")
        paper_text = st.text_area(
            "Question Paper / Reference Key:",
            value=default_paper,
            height=250,
            placeholder="Paste the original question paper or answer key here..."
        )
        
    with eval_col2:
        student_answers = st.text_area(
            "Student Submitted Answers:",
            height=250,
            placeholder="Paste the student's written or typed responses here (e.g. 1. A, 2. B, Short Ans 1: ...)..."
        )
        
    if st.button("🔍 Grade & Evaluate Student Answers"):
        if not api_key:
            st.error("Please provide a Gemini API Key in the sidebar.")
        elif not paper_text.strip():
            st.warning("Please provide the reference Question Paper.")
        elif not student_answers.strip():
            st.warning("Please paste the Student's Submitted Answers.")
        else:
            with st.spinner("Evaluating student response..."):
                evaluation_result = evaluate_student_answers(
                    api_key=api_key,
                    question_paper=paper_text,
                    student_answers=student_answers
                )
                
                st.markdown("### 🎯 Grading Report")
                st.markdown(evaluation_result)
