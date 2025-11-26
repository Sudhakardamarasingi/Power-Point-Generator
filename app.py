import io
import json
import requests
import streamlit as st
from pptx import Presentation

# Your n8n production webhook URL
WEBHOOK_URL = "https://sudha-mad-max-1997.app.n8n.cloud/webhook/f4892281-e1a0-429c-ae0a-16661a18e576"


def parse_ai_response(response: requests.Response) -> dict:
    """
    Handle different ways n8n might return the AI output.

    Cases:
    - Respond to Webhook JSON body = {{$json["output"]}}  -> response.json() is a string
    - Respond to Webhook JSON body = {{$json}}            -> response.json() is a dict
    """
    data = response.json()

    # If the HTTP body itself is just a JSON string that contains JSON
    if isinstance(data, str):
        return json.loads(data)

    # If the AI Agent output is nested under "output"
    if isinstance(data, dict) and "output" in data and isinstance(data["output"], str):
        return json.loads(data["output"])

    # Otherwise assume it is already the spec
    return data


def build_ppt_from_spec(spec: dict) -> bytes:
    """
    Build a .pptx from the JSON spec.

    Expected format:
    {
      "title": "Presentation Title",
      "slides": [
        {
          "heading": "Slide heading",
          "bullets": ["point 1", "point 2"],
          "notes": "optional presenter notes",
          "image_prompt": "optional"
        },
        ...
      ]
    }
    """
    prs = Presentation()

    slides = spec.get("slides", [])
    if not slides:
        raise ValueError("No slides found in AI response")

    # Title slide (first slide in list)
    first = slides[0]
    title_slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(title_slide_layout)
    title_shape = slide.shapes.title
    subtitle_shape = slide.placeholders[1]

    title_shape.text = spec.get("title") or first.get("heading", "Presentation")
    subtitle_shape.text = first.get("notes", "")

    # Content slides
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
            # First bullet as main text
            text_frame.text = bullets[0]
            # Remaining bullets
            for bullet in bullets[1:]:
                p = text_frame.add_paragraph()
                p.text = bullet
                p.level = 0

        # Notes and image_prompt are available in spec if you want to extend later

    buffer = io.BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


st.title("AGENTIC BASED POWER POINT GENERATOR")

prompt = st.text_area("Please write the details of how to create the PPT")

if st.button("Get PPT"):
    if not prompt.strip():
        st.warning("Please enter some details for the PPT.")
    else:
        with st.spinner("Generating your PPT..."):
            try:
                resp = requests.post(
                    WEBHOOK_URL,
                    json={"prompt": prompt},
                    timeout=180,
                )
            except requests.exceptions.RequestException as e:
                st.error(f"Failed to contact PPT generator:\n{e}")
            else:
                if resp.status_code != 200:
                    st.error(
                        f"PPT generation failed.\n"
                        f"Status: {resp.status_code}\n"
                        f"Body: {resp.text}"
                    )
                else:
                    try:
                        spec = parse_ai_response(resp)
                        ppt_bytes = build_ppt_from_spec(spec)
                    except Exception as e:
                        st.error(f"Error while building PPT from AI response:\n{e}")
                    else:
                        st.success("PPT generated successfully!")

                        st.download_button(
                            label="Download PPT",
                            data=ppt_bytes,
                            file_name="presentation.pptx",
                            mime=(
                                "application/vnd.openxmlformats-officedocument."
                                "presentationml.presentation"
                            ),
                        )
