import re
import io
import streamlit as st
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

st.title("Automated Resume Entity & Skill Extractor")

# ---------- Session state ----------
if "messages" not in st.session_state:
    st.session_state.messages = []
if "name" not in st.session_state:
    st.session_state.name = None
if "docx_bytes" not in st.session_state:
    st.session_state.docx_bytes = None  # initialize here

# ---------- Intro (only first time) ----------
if not st.session_state.messages:
    st.session_state.messages.append({
        "role": "assistant",
        "content": (
            "Hello! I am Automated Resume Entity & Skill Extractor.\n\n"
            "I will ask you a series of questions to build a resume automatically."
        )
    })

# ---------- Helper: extract name from sentence ----------
def extract_name(text: str) -> str:
    text = text.strip()

    patterns = [
        r"my name is\s+(.+)",   # My name is Zay Yar Min.
        r"i am\s+(.+)",         # I am Zay Yar Min.
        r"i'm\s+(.+)"           # I'm Zay Yar Min.
    ]

    for pat in patterns:
        m = re.search(pat, text, flags=re.IGNORECASE)
        if m:
            name = m.group(1).strip(". ").strip()
            return name

    # Fallback: use full text
    return text

# ---------- Render existing history once ----------
for msg in st.session_state.messages:
    with st.chat_message(msg["role"]):
        st.write(msg["content"])

# ---------- Ask for name if not set ----------
if st.session_state.name is None:
    # Add the question as a normal bot message only once
    if not any("please tell me your full name" in m["content"].lower()
               for m in st.session_state.messages):
        question = "First, please tell me your full name."
        st.session_state.messages.append({
            "role": "assistant",
            "content": question
        })
        with st.chat_message("assistant"):
            st.write(question)

# ---------- Input box ----------
user_input = st.chat_input("Type your answer here...")

if user_input:
    # Store user message
    st.session_state.messages.append({
        "role": "user",
        "content": user_input
    })

    # Process name only if we don't have it yet
    if st.session_state.name is None:
        name = extract_name(user_input)
        st.session_state.name = name

        bot_reply = (
            f"Nice to meet you, {name}! "
            f"I will use this as the name at the top of your resume."
        )

        st.session_state.messages.append({
            "role": "assistant",
            "content": bot_reply
        })

        # Show only the new messages for this turn
        with st.chat_message("user"):
            st.write(user_input)
        with st.chat_message("assistant"):
            st.write(bot_reply)

# ---------- If we have name, generate Word file (test) ----------
if st.session_state.name is not None and st.session_state.docx_bytes is None:
    with st.chat_message("assistant"):
        st.write(
            "Please wait, I am generating your resume Word file using your name..."
        )

    progress_bar = st.progress(0)
    for pct in range(0, 101, 25):
        progress_bar.progress(pct)

    # Create Word document in memory
    doc = Document()  # [web:70][web:67]

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(st.session_state.name)
    run.bold = True
    
    run.font.size = Pt(36)
    run.font.name = "Times New Roman"

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    st.session_state.docx_bytes = buffer.read()

    with st.chat_message("assistant"):
        st.write(
            "Your test resume Word file is ready. You can download it below."
        )

# ---------- Download button ----------
if st.session_state.docx_bytes is not None:
    st.download_button(
        label="Download Resume (test)",
        data=st.session_state.docx_bytes,
        file_name="resume_test.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )