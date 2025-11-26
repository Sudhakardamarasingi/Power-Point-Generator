import io
import json
import time
import requests
import streamlit as st
from pptx import Presentation

# ---------- CONFIG ----------
WEBHOOK_URL = "https://sudha-mad-max-1997.app.n8n.cloud/webhook/f4892281-e1a0-429c-ae0a-16661a18e576"

st.set_page_config(
    page_title="Agentic PowerPoint Designer",
    page_icon="🧩",
    layout="wide",
)

# ---------- STYLES ----------
st.markdown(
    """
    <style>
    .block-container {
        max-width: 880px;
        padding-top: 2rem;
        margin-left: auto;
        margin-right: auto;
    }
    .stApp {
        background: radial-gradient(circle at top left, #2b3a67 0, #050816 45%, #050816 100%);
        color: #f9fafb;
    }
    .glass-card {
        background: rgba(15, 23, 42, 0.92);
        border-radius: 18px;
        padding: 2rem;
        border: 1px solid rgba(255,255,255,0.08);
        box-shadow: 0 18px 50px rgba(0,0,0,0.45);
        backdrop-filter: blur(16px);
    }
    textarea {
        border-radius: 14px !important;
        border: 1px solid rgba(148, 163, 184, 0.6) !important;
    }
    button[kind="primary"] {
        border-radius: 999px !important;
        font-weight: 600 !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ---------- HELPERS ----------
def parse_ai_response(response: requests.Response) -> dict:
    raw = response.text.strip()
    if not raw:
        raise ValueError("AI response is empty")
    try:
        return json.loads(raw)
    except Exception:
        st.error("AI returned invalid JSON:")
        st.code(raw)
        raise

def build_ppt_from_spec(spec: dict) -> bytes:
    prs = Presentation()
    slides = spec.get("slides", [])

    first = slides[0]
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = spec.get("title", "")
    slide.placeholders[1].text = first.get("notes", "")

    for slide_spec in slides[1:]:
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = slide_spec.get("heading", "")
        tf = slide.placeholders[1].text_frame
        tf.clear()

        bullets = slide_spec.get("bullets", [])
        if bullets:
            tf.text = bullets[0]
            for b in bullets[1:]:
                p = tf.add_paragraph()
                p.text = b
                p.level = 0

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ---------- STATE ----------
if "ppt_bytes" not in st.session_state:
    st.session_state["ppt_bytes"] = None

# ---------- HEADER ----------
st.markdown(
    """
    <div style="text-align:center; margin-bottom:2.5rem;">
        <h1 style="margin-bottom:0.3rem;">Agentic PowerPoint Designer</h1>
        <p style="opacity:0.85;">Turn a single prompt into a complete, editable slide deck.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

# ---------- MAIN CARD ----------
with st.container():
    st.markdown('<div class="glass-card">', unsafe_allow_html=True)

    st.subheader("Describe your presentation", anchor=False)
    st.caption("Enter the topic and any key points you want covered.")

    prompt = st.text_area(
        label="",
        placeholder="Example: Create an 8-slide PPT on agentic AI...",
        height=160,
    )

    generate = st.button("✨ Generate PPT", use_container_width=True)

    progress_placeholder = st.empty()
    status_placeholder = st.empty()
    success_placeholder = st.empty()
    download_placeholder = st.empty()

    st.markdown("</div>", unsafe_allow_html=True)


# ---------- ACTION ----------
if generate:
    if not prompt.strip():
        st.warning("Please enter a topic.")
    else:
        st.session_state["ppt_bytes"] = None
        prog = progress_placeholder.progress(0)

        status_placeholder.write("🧩 Preparing your request...")
        for i in range(10):
            prog.progress(i * 5)
            time.sleep(0.03)

        status_placeholder.write("🔗 Contacting n8n agent...")
        resp = requests.post(WEBHOOK_URL, json={"prompt": prompt})

        if resp.status_code != 200:
            st.error("Backend error")
        else:
            status_placeholder.write("🤖 Agentic AI finished. Parsing JSON...")
            spec = parse_ai_response(resp)

            status_placeholder.write("📑 Building PPT...")
            ppt_bytes = build_ppt_from_spec(spec)
            st.session_state["ppt_bytes"] = ppt_bytes

            prog.progress(100)
            success_placeholder.success("PPT generated successfully.")


# ---------- DOWNLOAD ----------
if st.session_state["ppt_bytes"]:
    download_placeholder.download_button(
        label="⬇️ Download PPTX",
        data=st.session_state["ppt_bytes"],
        file_name="presentation.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        use_container_width=True,
    )
