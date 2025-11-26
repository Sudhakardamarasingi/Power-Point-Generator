import io
import json
import time
import requests
import streamlit as st
from pptx import Presentation

WEBHOOK_URL = "https://sudha-mad-max-1997.app.n8n.cloud/webhook/f4892281-e1a0-429c-ae0a-16661a18e576"

st.set_page_config(
    page_title="Agentic PPT Designer",
    page_icon="🎨",
    layout="wide",
)

# ---------- STYLES ----------
st.markdown(
    """
    <style>
    /* Center the whole app content */
    .block-container {
        max-width: 880px;
        padding-top: 2.5rem;
        padding-bottom: 3rem;
        margin: 0 auto;
    }

    .stApp {
        background: radial-gradient(circle at top left, #2b3a67 0, #050816 45%, #050816 100%);
        color: #f9fafb;
        font-family: system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
    }

    .glass-card {
        background: rgba(15, 23, 42, 0.9);
        border-radius: 18px;
        padding: 1.6rem 1.8rem 1.8rem;
        border: 1px solid rgba(148, 163, 184, 0.35);
        box-shadow: 0 18px 50px rgba(15, 23, 42, 0.8);
        backdrop-filter: blur(16px);
    }

    .pill-badge {
        display: inline-flex;
        align-items: center;
        gap: 0.35rem;
        border-radius: 999px;
        border: 1px solid rgba(148, 163, 184, 0.55);
        padding: 0.2rem 0.65rem;
        font-size: 0.75rem;
        opacity: 0.95;
    }

    .pill-dot {
        width: 8px;
        height: 8px;
        border-radius: 999px;
        background: linear-gradient(135deg,#22c55e,#a855f7);
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
    title_slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(title_slide_layout)
    title_shape = slide.shapes.title
    subtitle_shape = slide.placeholders[1]

    title_shape.text = spec.get("title") or first.get("heading", "Presentation")
    subtitle_shape.text = first.get("notes", "")

    content_layout = prs.slide_layouts[1]

    for slide_spec in slides[1:]:
        slide = prs.slides.add_slide(content_layout)
        title_shape = slide.shapes.title
        body_shape = slide.placeholders[1]

        title_shape.text = slide_spec.get("heading", "")

        text_frame = body_shape.text_frame
        text_frame.clear()

        bullets = slide_spec.get("bullets", [])
        if bullets:
            text_frame.text = bullets[0]
            for bullet in bullets[1:]:
                p = text_frame.add_paragraph()
                p.text = bullet
                p.level = 0

    buffer = io.BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()

# ---------- STATE ----------
if "ppt_bytes" not in st.session_state:
    st.session_state["ppt_bytes"] = None

# ---------- HEADER ----------
st.markdown(
    """
    <div style="text-align:center; margin-bottom:1.8rem;">
      <div class="pill-badge">
        <span class="pill-dot"></span>
        <span>Agentic AI · n8n · Streamlit</span>
      </div>
      <h1 style="margin-top:1rem; margin-bottom:0.3rem;">
        Agentic PowerPoint Designer
      </h1>
      <p style="opacity:0.85; max-width:540px; margin:0 auto;">
        Turn a single prompt into a complete, editable slide deck.
      </p>
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
        placeholder="Example: Create an 8-slide PPT explaining agentic AI for non-technical business leaders...",
        height=160,
    )

    generate = st.button("✨ Generate PPT", use_container_width=True)

    # Progress + status live inside the same card
    progress_placeholder = st.empty()
    status_placeholder = st.empty()

    # Download sits inside the card too
    download_placeholder = st.empty()

    st.markdown("</div>", unsafe_allow_html=True)

# ---------- ACTION + SMOOTH PROGRESS ----------
if generate:
    if not prompt or not prompt.strip():
        st.warning("Please enter a topic or description for the PPT.")
    else:
        progress = progress_placeholder.progress(0)
        status = status_placeholder

        try:
            status.markdown("🧩 Preparing your request...")
            for i in range(1, 16):
                progress.progress(i)
                time.sleep(0.02)

            status.markdown("🔗 Contacting agentic backend...")
            progress.progress(25)
            resp = requests.post(
                WEBHOOK_URL,
                json={"prompt": prompt},
                timeout=180,
            )

            status.markdown("🤖 Agent finished. Parsing slide structure...")
            for i in range(26, 61):
                progress.progress(i)
                time.sleep(0.01)

            if resp.status_code != 200:
                progress.progress(0)
                status_placeholder.empty()
                st.error(
                    f"PPT generation failed.\n"
                    f"Status: {resp.status_code}\n"
                    f"Body: {resp.text}"
                )
            else:
                spec = parse_ai_response(resp)

                status.markdown("📑 Building your PowerPoint file...")
                for i in range(61, 90):
                    progress.progress(i)
                    time.sleep(0.01)

                ppt_bytes = build_ppt_from_spec(spec)
                st.session_state["ppt_bytes"] = ppt_bytes

                progress.progress(100)
                status.markdown("✅ Deck ready to download.")
                st.success("PPT generated successfully.")

        except Exception as e:
            progress.progress(0)
            status_placeholder.empty()
            st.error(f"Error while generating PPT:\n{e}")

# ---------- DOWNLOAD (inside card, under progress) ----------
if st.session_state["ppt_bytes"]:
    with download_placeholder:
        st.markdown("---")
        st.download_button(
            label="⬇️ Download PPTX",
            data=st.session_state["ppt_bytes"],
            file_name="presentation.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True,
        )
