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

    /* Center button containers and keep width to content */
    div.stButton, div.stDownloadButton {
        text-align: center;
    }

    /* Purple primary button */
    div.stButton > button {
        background: #8b5cf6;
        color: #ffffff;
        border-radius: 999px;
        border: none;
        padding: 0.55rem 1.6rem;
        font-weight: 600;
        font-size: 0.95rem;
        box-shadow: 0 10px 25px rgba(139, 92, 246, 0.35);
        transition: all 0.15s ease-out;
        width: auto;
    }
    div.stButton > button:hover {
        background: #7c3aed;
        transform: translateY(-1px);
        box-shadow: 0 14px 30px rgba(139, 92, 246, 0.45);
    }

    /* Cyan download pill */
    div.stDownloadButton > button {
        background: #06b6d4;
        color: #0f172a;
        border-radius: 999px;
        border: none;
        padding: 0.6rem 1.8rem;
        font-weight: 600;
        font-size: 0.95rem;
        box-shadow: 0 10px 25px rgba(6, 182, 212, 0.45);
        transition: all 0.15s ease-out;
        width: auto;
    }
    div.stDownloadButton > button:hover {
        background: #0891b2;
        transform: translateY(-2px);
        box-shadow: 0 14px 30px rgba(6, 182, 212, 0.6);
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ---------- HELPERS ----------
def parse_ai_response(response: requests.Response):
    """
    Parse AI/n8n response.
    - Try direct JSON.
    - If that fails, try to extract the first {...} block and parse that.
    - On failure, show an error and return None.
    """
    raw = response.text.strip()

    if not raw:
        st.error(
            "Backend returned an empty response. "
            "Check your n8n 'Respond to Webhook' node and make sure it returns the AI JSON."
        )
        return None

    # 1) Try direct JSON parse first
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        pass  # we'll try to clean it up below

    # 2) Try to extract the first JSON object between outermost { ... }
    start = raw.find("{")
    end = raw.rfind("}")
    if start != -1 and end != -1 and end > start:
        candidate = raw[start : end + 1]
        try:
            return json.loads(candidate)
        except json.JSONDecodeError as e2:
            st.error(f"AI returned invalid JSON even after cleanup: {e2}")
            st.code(candidate[:1000])
            return None

    # 3) If we can't even find a { ... } block, show the raw text
    st.error("AI returned text that doesn't look like JSON.")
    st.code(raw[:1000])
    return None


def build_ppt_from_spec(spec: dict) -> bytes:
    prs = Presentation()
    slides = spec.get("slides", [])
    if not slides:
        raise ValueError("No 'slides' array found in AI response")

    # Title slide
    first = slides[0]
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = spec.get("title", first.get("heading", "Presentation"))
    slide.placeholders[1].text = first.get("notes", "")

    # Content slides
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
for key, default in [
    ("ppt_bytes", None),
    ("ready", False),
    ("is_generating", False),
    ("trigger_generate", False),
]:
    if key not in st.session_state:
        st.session_state[key] = default


def start_generation():
    """Callback when user clicks 'Create a presentation with AI'."""
    st.session_state["trigger_generate"] = True
    st.session_state["is_generating"] = True
    st.session_state["ready"] = False
    st.session_state["ppt_bytes"] = None


def clear_after_download():
    """After download, hide download button and show generate button again."""
    st.session_state["ppt_bytes"] = None
    st.session_state["ready"] = False
    st.session_state["is_generating"] = False
    st.session_state["trigger_generate"] = False


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

    prompt = st.text_area(
        label="",
        key="prompt_text",  # store in session_state
        placeholder="Example: Create an 8-slide PPT on agentic AI for non-technical business leaders...",
        height=160,
    )

    # Only show the generate button if we're NOT generating and NOT ready
    show_generate_button = (not st.session_state["is_generating"]) and (not st.session_state["ready"])
    if show_generate_button:
        st.button("Create a presentation with AI", on_click=start_generation)

    progress_placeholder = st.empty()
    status_placeholder = st.empty()
    success_placeholder = st.empty()

    # If trigger_generate is set, run the full generation pipeline
    if st.session_state["trigger_generate"]:
        # Basic validation
        if not st.session_state.get("prompt_text", "").strip():
            st.warning("Please enter a topic.")
            st.session_state["is_generating"] = False
            st.session_state["trigger_generate"] = False
        else:
            prog = progress_placeholder.progress(0)
            status_placeholder.write("🧩 Preparing your request...")

            for i in range(1, 16):
                prog.progress(i)
                time.sleep(0.02)

            status_placeholder.write("🔗 Contacting n8n agent...")
            try:
                resp = requests.post(
                    WEBHOOK_URL,
                    json={"prompt": st.session_state["prompt_text"]},
                    timeout=180,
                )
            except Exception as e:
                prog.progress(0)
                status_placeholder.empty()
                st.error(f"Error talking to backend: {e}")
                st.session_state["is_generating"] = False
                st.session_state["trigger_generate"] = False
            else:
                if resp.status_code != 200:
                    prog.progress(0)
                    status_placeholder.empty()
                    st.error(f"Backend error: {resp.status_code}\n{resp.text}")
                    st.session_state["is_generating"] = False
                    st.session_state["trigger_generate"] = False
                else:
                    status_placeholder.write("🤖 Agent finished. Parsing JSON...")
                    spec = parse_ai_response(resp)

                    if spec is None:
                        prog.progress(0)
                        st.session_state["is_generating"] = False
                        st.session_state["trigger_generate"] = False
                    else:
                        status_placeholder.write("📑 Building PPT...")
                        for i in range(16, 95):
                            prog.progress(i)
                            time.sleep(0.01)

                        try:
                            ppt_bytes = build_ppt_from_spec(spec)
                        except Exception as e:
                            prog.progress(0)
                            status_placeholder.empty()
                            st.error(f"Error while building PPT: {e}")
                            st.session_state["ppt_bytes"] = None
                            st.session_state["ready"] = False
                        else:
                            st.session_state["ppt_bytes"] = ppt_bytes
                            st.session_state["ready"] = True
                            prog.progress(100)
                            status_placeholder.write("✅ Presentation ready to download.")
                            success_placeholder.success("PPT generated successfully.")

                        st.session_state["is_generating"] = False
                        st.session_state["trigger_generate"] = False

    # ---------- DOWNLOAD BUTTON (INSIDE CARD) ----------
    if st.session_state["ready"] and st.session_state["ppt_bytes"]:
        st.markdown("<br>", unsafe_allow_html=True)
        st.download_button(
            label="⬇ Download PPTX",
            data=st.session_state["ppt_bytes"],
            file_name="presentation.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            key="download_ppt",
            on_click=clear_after_download,
        )
