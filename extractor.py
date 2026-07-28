import re
import io
import streamlit as st
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

st.title("Automated Resume Entity & Skill Extractor")

# ---------- Session states ----------
if "messages" not in st.session_state:
    st.session_state.messages = []
if "name" not in st.session_state:
    st.session_state.name = None
if "DOB" not in st.session_state:
    st.session_state.DOB = None
if "ADDR" not in st.session_state:
    st.session_state.ADDR = None
if "docx_bytes" not in st.session_state:
    st.session_state.docx_bytes = None
if "step" not in st.session_state:
    # steps: "intro", "name", "dob", "done"
    st.session_state.step = "intro"

# ---------- Intro (only first time) ----------
if not st.session_state.messages:
    st.session_state.messages.append({
        "role": "assistant",
        "content": (
            "မင်္ဂလာပါ၊ ကျွန်တော်ကတော့ အလိုအလျောက် CV နဲ့ ကျွမ်းကျင်မှုတွေ ထုတ်ဖော်ရေးသားပေးသွားမှာပဲဖြစ်ပါတယ်။\n\n"
            "CV ရေးသားဖို့ ဘယ်ကနေ ဘယ်လိုစရေးရမှန်း မသိတဲ့ လူတွေအတွက် အကူအညီပေးဖို့ ရည်ရွယ်ပါတယ်။ \n CV ကို တည်ဆောက်ဖို့အတွက် ကျွန်တော်က မေးခွန်းအချို့ကို မေးမြန်းသွားမှာပဲ ဖြစ်ပါတယ်။"
        )
    })
    st.session_state.step = "name"  # after intro, go to name question

# ---------- Helper methods ----------
def extract_name(text: str) -> str:
    if not text:
        return ""
    text = text.strip()

    patterns = [
        r"my name is\s+(.+)",   # My name is ...
        r"i am\s+(.+)",         # I am ...
        r"i'm\s+(.+)",           # I'm ...
        r".*ကတော့\s+(.+)\s+.*(ဖြစ်ပါတယ်။|ပါ။|ဖြစ်တယ်။)$",
        r".*က\s+(.+)\s+.*(ဖြစ်ပါတယ်။|ပါ။|ဖြစ်တယ်။)$"
    ]

    for pat in patterns:
        m = re.search(pat, text, flags=re.IGNORECASE)
        if m:
            name = m.group(1).strip(". ").strip()
            return name

    return text.strip(". ")  # fallback: full text

def extract_dob(text: str) -> str:
    if not text:
        return ""
    text = text.strip(". ").strip()

    patterns = [
        r".*\sis\s+(0[1-9]|[12][0-9]|3[0-1])\/(0[1-9]|1[0-2])\/([0-9]{4})$",
        r"(0[1-9]|[12][0-9]|3[0-1])\/(0[1-9]|1[0-2])\/([0-9]{4})$",
        r".*ကတော့\s+(0[1-9]|[12][0-9]|3[0-1])\/(0[1-9]|1[0-2])\/([0-9]{4})\s+.*(ဖြစ်ပါတယ်။|ပါ။|ဖြစ်တယ်။)$",
        r".*က\s+(0[1-9]|[12][0-9]|3[0-1])\/(0[1-9]|1[0-2])\/([0-9]{4})\s+.*(ဖြစ်ပါတယ်။|ပါ။|ဖြစ်တယ်။)$"

    ]

    for pat in patterns:
        m = re.search(pat, text, flags=re.IGNORECASE)
        if m:
            day = m.group(1).strip()
            month = m.group(2).strip()
            year = m.group(3).strip()

            match month:
                case "01": mon = "January"
                case "02": mon = "February"
                case "03": mon = "March"
                case "04": mon = "April"
                case "05": mon = "May"
                case "06": mon = "June"
                case "07": mon = "July"
                case "08": mon = "August"
                case "09": mon = "September"
                case "10": mon = "October"
                case "11": mon = "November"
                case "12": mon = "December"
                case _: mon = "Unknown Month"

            return f"{day} {mon} {year}"

    return text  # fallback

def extract_addr(text: str) -> str:
    if not text:
        return ""
    text = text.strip(". ").strip()

    patterns = [
        r".*\sis\s+(.+)\.$",
        r".*ကတော့\s+(.+)\s+.*(ဖြစ်ပါတယ်။|ပါ။|ဖြစ်တယ်။)$",
        r".*က\s+(.+)\s+.*(ဖြစ်ပါတယ်။|ပါ။|ဖြစ်တယ်။)$",
        r".*ကတော့\s+(.+)\s+မှာ.*",
        r".*က\s+(.+)\s+မှာ.*"
    ]

    for pat in patterns:
        m = re.search(pat, text, flags=re.IGNORECASE)
        if m:
            addr = m.group(1).strip()

            return addr

    return text  # fallback


# ---------- Render existing history ----------
for msg in st.session_state.messages:
    with st.chat_message(msg["role"]):
        st.write(msg["content"])

# ---------- Name step ----------
if st.session_state.step == "name" and st.session_state.name is None:
    # Ask only once
    question = "ပထမဆုံး သင့်ရဲ့ နာမည်အပြည့်အစုံကို အင်္ဂလိပ်လို ရေးပေးပါ။"

    # Ask only once: check for exact question text in history
    if not any(
        m["role"] == "assistant" and m["content"] == question
        for m in st.session_state.messages
    ):
        st.session_state.messages.append({
            "role": "assistant",
            "content": question
        })
        with st.chat_message("assistant"):
            st.write(question)

    user_input = st.chat_input("သင့်ရဲ့ နာမည်ကို ဒီမှာရေးပေးပါ...")

    if user_input:
        # Store user message
        st.session_state.messages.append({
            "role": "user",
            "content": user_input
        })

        # Process name
        name = extract_name(user_input)
        st.session_state.name = name

        with st.chat_message("user"):
            st.write(user_input)

        # Move to DOB step
        st.session_state.step = "dob"
        st.rerun()

# ---------- DOB step ----------
elif st.session_state.step == "dob" and st.session_state.DOB is None:
    dobQuestion = "အခု သင့်ရဲ့ မွေးသက္ကရာဇ်ကို (dd/mm/yyyy) ပုံစံဖြင့် အင်္ဂလိပ် ဂဏန်းများ ထည့်သွင်းပေးပါ။"

    # Ask only once: check for exact question text in history
    if not any(
        m["role"] == "assistant" and m["content"] == dobQuestion
        for m in st.session_state.messages
    ):
        st.session_state.messages.append({
            "role": "assistant",
            "content": dobQuestion
        })
        with st.chat_message("assistant"):
            st.write(dobQuestion)

    dobInput = st.chat_input("သင့်ရဲ့ မွေးသက္ကရာဇ်ကို ဒီမှာ ရေးပေးပါ...")

    if dobInput:
        st.session_state.messages.append({
            "role": "user",
            "content": dobInput
        })

        dob = extract_dob(dobInput)
        st.session_state.DOB = dob

        with st.chat_message("user"):
            st.write(dobInput)

        # All basic info collected
        st.session_state.step = "addr"
        st.rerun()

# ---------- Addr step ----------
elif st.session_state.step == "addr" and st.session_state.ADDR is None:
    addrQuestion = "အခု သင့်ရဲ့ နေရပ်လိပ်စာကို ထည့်သွင်းပေးပါ။"

    # Ask only once: check for exact question text in history
    if not any(
        m["role"] == "assistant" and m["content"] == addrQuestion
        for m in st.session_state.messages
    ):
        st.session_state.messages.append({
            "role": "assistant",
            "content": addrQuestion
        })
        with st.chat_message("assistant"):
            st.write(addrQuestion)

    addrInput = st.chat_input("သင့်ရဲ့ နေရပ်လိပ်စာကို ဒီမှာ ရေးပေးပါ...")

    if addrInput:
        st.session_state.messages.append({
            "role": "user",
            "content": addrInput
        })

        addr = extract_addr(addrInput)
        st.session_state.ADDR = addr

        with st.chat_message("user"):
            st.write(addrInput)

        # All basic info collected
        st.session_state.step = "done"
        st.rerun()

# ---------- Generate Word file ----------
if st.session_state.step == "done" and st.session_state.docx_bytes is None:
    with st.chat_message("assistant"):
        st.write(
            "Please wait, I am generating your resume Word file using your name..."
        )

    progress_bar = st.progress(0)
    for pct in range(0, 101, 25):
        progress_bar.progress(pct)

    doc = Document()

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    titleRun = title.add_run("Curriculum Vitae")
    titleRun.bold = True
    titleRun.underline = True
    titleRun.font.size = Pt(36)
    titleRun.font.name = "Times New Roman"

    nameP = doc.add_paragraph()
    nameRun = nameP.add_run(st.session_state.name)
    nameRun.bold = True
    nameRun.font.size = Pt(24)
    nameRun.font.name = "Times New Roman"

    dobP = doc.add_paragraph()
    dobBullet = dobP.add_run("\u2666")
    dobBullet.bold = True
    dobBullet.font.size = Pt(12)
    dobBullet.font.name = "Segoe UI Symbol"
    dobLable = dobP.add_run("Date of Birth\t: ")
    dobLable.bold = True
    dobLable.font.size = Pt(12)
    dobLable.font.name = "Times New Roman"
    dobValue = dobP.add_run(st.session_state.DOB)    
    dobValue.font.size = Pt(12)
    dobValue.font.name = "Times New Roman"

    addrP = doc.add_paragraph()
    addrBullet = addrP.add_run("\u2666")
    addrBullet.bold = True
    addrBullet.font.size = Pt(12)
    addrBullet.font.name = "Segoe UI Symbol"
    addrLable = addrP.add_run("Address\t\t: ")
    addrLable.bold = True
    addrLable.font.size = Pt(12)
    addrLable.font.name = "Times New Roman"
    addrValue = addrP.add_run(st.session_state.ADDR)    
    addrValue.font.size = Pt(12)
    addrValue.font.name = "Times New Roman"

        
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
        label="Download CV",
        data=st.session_state.docx_bytes,
        file_name="cv.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )