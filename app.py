import io
import json
import time
import requests
import streamlit as st
from pptx import Presentation

WEBHOOK_URL = "https://sudha-mad-max-1997.app.n8n.cloud/webhook/f4892281-e1a0-429c-ae0a-16661a18e576"

st.set_page_config(
    page_title="Agentic PowerPoint Designer",
    page_icon="🧩",
    layout="wide",
)

# ---------- GLOBAL STYLES ----------
st.markdown(
    """
    <style>
    .stApp {
        background: radial-gradient(circle at top left, #2b3a67 0, #050816 45%, #050816 100%);
        color: #f9fafb;
        font-family: system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
    }
    .block-container {
        max-width: 880px;
        padding-top: 2rem;
        padding-bottom: 3rem;
        margin-left: auto;
        margin-right: auto;
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
    return json.loads(raw)

def build_ppt_from_spec(spec: dict) -> bytes:
    prs = Presentation()
    slides = spec.get("slides", [])
    if not slides:
        raise ValueError("No 'slides' array found in AI response")

    first = slides[0]
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = spec.get("title", first.get("heading", "Presentation"))
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

# ---------- SESSION STATE ----------
if "ppt_bytes" not in st.session_state:
    st.session_state["ppt_bytes"] = None
if "has_generated" not in st.session_state:
    st.session_state["has_generated"] = False

# ---------- HEADER ----------
st.markdown(
    """
    <div style="text-align:center; margin-bottom:2rem;">
      <h1 style="margin-bottom:0.3rem;">Agentic PowerPoint Designer</h1>
      <p style="opacity:0.85;">Turn a single prompt into a complete, editable slide deck.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

# ---------- MAIN CARD ----------
with st.container(border=True):
    st.subheader("Describe your presentation", anchor=False)
    st.caption("Enter the topic and any key points you want covered.")

    # This is the ONLY prompt area
    prompt = st.text_area(
        label="",
        placeholder="Example: Create an 8-slide PPT on agentic AI for non-technical business leaders...",
        height=160,
    )

    generate = st.button("✨ Generate PPT", use_container_width=True)

    progress_placeholder = st.empty()
    status_placeholder = st.empty()
    success_placeholder = st.empty()
    download_placeholder = st.empty()

    # ---------- ACTION ----------
    if generate:
        if not prompt.strip():
            st.warning("Please enter a topic.")
            st.session_state["has_generated"] = False
            st.session_state["ppt_bytes"] = None
        else:
            st.session_state["has_generated"] = False
            st.session_state["ppt_bytes"] = None

            prog = progress_placeholder.progress(0)
            status_placeholder.write("🧩 Preparing your request...")
            for i in range(1, 16):
                prog.progress(i)
                time.sleep(0.02)

            status_placeholder.write("🔗 Contacting n8n agent...")
            try:
                resp = requests.post(WEBHOOK_URL, json={"prompt": prompt}, timeout=180)
            except Exception as e:
                prog.progress(0)
                status_placeholder.empty()
                st.error(f"Error talking to backend: {e}")
            else:
                if resp.status_code != 200:
                    prog.progress(0)
                    status_placeholder.empty()
                    st.error(f"Backend error: {resp.status_code}\n{resp.text}")
                else:
                    status_placeholder.write("🤖 Agent finished. Parsing JSON...")
                    spec = parse_ai_response(resp)

                    status_placeholder.write("📑 Building PPT...")
                    for i in range(16, 95):
                        prog.progress(i)
                        time.sleep(0.01)

                    ppt_bytes = build_ppt_from_spec(spec)
                    st.session_state["ppt_bytes"] = ppt_bytes
                    st.session_state["has_generated"] = True

                    prog.progress(100)
                    success_placeholder.success("PPT generated successfully.")

    # ---------- DOWNLOAD (inside the same card, only after generation) ----------
    if st.session_state["has_generated"] and st.session_state["ppt_bytes"]:
        download_placeholder.download_button(
            label="⬇️ Download PPTX",
            data=st.session_state["ppt_bytes"],
            file_name="presentation.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True,
        )
