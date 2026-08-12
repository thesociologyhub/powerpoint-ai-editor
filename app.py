from flask import Flask, request, send_file, jsonify
from pptx import Presentation
import io
import os
import requests

app = Flask(__name__)

OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY")

SYSTEM_PROMPT = """
You are The Sociology Hub PowerPoint Proofreading Assistant.

Your job is to make ONLY safe editorial corrections.

You MAY correct:
- spelling
- grammar
- punctuation
- obvious typos
- accidental duplicated words
- capitalisation consistency

You MUST NOT change:
- sociology content
- sociological terminology
- curriculum content
- teaching activities
- exam questions
- learning objectives
- pedagogical decisions
- factual claims
- examples
- meaning or tone

Preserve the wording as closely as possible.

Return ONLY the corrected text.
Do not explain your changes.
"""

def correct_text(text):
    if not text or not text.strip():
        return text

    response = requests.post(
        "https://api.openai.com/v1/responses",
        headers={
            "Authorization": f"Bearer {OPENAI_API_KEY}",
            "Content-Type": "application/json",
        },
        json={
            "model": "gpt-5-mini",
            "instructions": SYSTEM_PROMPT,
            "input": text,
        },
        timeout=120,
    )

    response.raise_for_status()
    data = response.json()

    for item in data.get("output", []):
        for content in item.get("content", []):
            if content.get("type") == "output_text":
                return content.get("text", text)

    return text


@app.route("/")
def home():
    return "PowerPoint AI Editor is running!"


@app.route("/edit", methods=["POST"])
def edit_powerpoint():
    if "file" not in request.files:
        return jsonify({"error": "No PowerPoint file received"}), 400

    if not OPENAI_API_KEY:
        return jsonify({"error": "OPENAI_API_KEY is not configured"}), 500

    uploaded_file = request.files["file"]

    presentation = Presentation(uploaded_file)

    for slide in presentation.slides:
        for shape in slide.shapes:
            if not hasattr(shape, "text_frame"):
                continue

            for paragraph in shape.text_frame.paragraphs:
                original_text = "".join(run.text for run in paragraph.runs)

                if not original_text.strip():
                    continue

                corrected_text = correct_text(original_text)

                if corrected_text != original_text and paragraph.runs:
                    paragraph.runs[0].text = corrected_text

                    for run in paragraph.runs[1:]:
                        run.text = ""

    output = io.BytesIO()
    presentation.save(output)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="corrected-powerpoint.pptx",
        mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )
