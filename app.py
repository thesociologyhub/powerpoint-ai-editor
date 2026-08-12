from flask import Flask, request, send_file, jsonify
from openai import OpenAI
from lxml import etree
import tempfile
import zipfile
import os
import json
import shutil

app = Flask(__name__)

client = OpenAI(
    api_key=os.environ.get("OPENAI_API_KEY"),
    timeout=60.0,
    max_retries=2,
)

SYSTEM_PROMPT = """
You are The Sociology Hub PowerPoint Proofreading Assistant.

Make ONLY safe editorial corrections.

You MAY correct:
- spelling
- grammar
- punctuation
- obvious typos
- duplicated words
- capitalisation consistency

Use UK English.

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

Preserve wording as closely as possible.

Return ONLY valid JSON in this exact format:
{"items":[{"id":0,"text":"corrected text"}]}

Return every supplied id exactly once and in the same order.
"""

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main"
}


def proofread_texts(items):
    response = client.responses.create(
        model="gpt-5-mini",
        instructions=SYSTEM_PROMPT,
        input=json.dumps(items, ensure_ascii=False),
    )

    result = json.loads(response.output_text)
    return result["items"]


def edit_slide_xml(xml_bytes):
    root = etree.fromstring(xml_bytes)

    text_nodes = root.xpath("//a:t", namespaces=NS)

    items = []

    for index, node in enumerate(text_nodes):
        if node.text and node.text.strip():
            items.append({
                "id": index,
                "text": node.text
            })

    if not items:
        return xml_bytes

    batch_size = 30

    for start in range(0, len(items), batch_size):
        batch = items[start:start + batch_size]
        corrected_items = proofread_texts(batch)

        correction_map = {
            item["id"]: item["text"]
            for item in corrected_items
        }

        for item in batch:
            corrected = correction_map.get(item["id"])

            if corrected is not None:
                text_nodes[item["id"]].text = corrected

    return etree.tostring(
        root,
        xml_declaration=True,
        encoding="UTF-8",
        standalone=True
    )


@app.route("/")
def home():
    return "PowerPoint AI Editor is running!"


@app.route("/edit", methods=["POST"])
def edit_powerpoint():
    input_path = None
    output_path = None

    try:
        if "file" not in request.files:
            return jsonify({"error": "No PowerPoint file received"}), 400

        if not os.environ.get("OPENAI_API_KEY"):
            return jsonify({"error": "OPENAI_API_KEY is not configured"}), 500

        uploaded_file = request.files["file"]

        with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as temp_input:
            uploaded_file.save(temp_input.name)
            input_path = temp_input.name

        with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as temp_output:
            output_path = temp_output.name

        with zipfile.ZipFile(input_path, "r") as source_zip:
            with zipfile.ZipFile(
                output_path,
                "w",
                compression=zipfile.ZIP_DEFLATED
            ) as target_zip:

                for item in source_zip.infolist():
                    file_data = source_zip.read(item.filename)

                    if (
                        item.filename.startswith("ppt/slides/slide")
                        and item.filename.endswith(".xml")
                    ):
                        file_data = edit_slide_xml(file_data)

                    target_zip.writestr(item, file_data)

        original_name = uploaded_file.filename or "presentation.pptx"

        if original_name.lower().endswith(".pptx"):
            new_name = original_name[:-5] + " - Corrected.pptx"
        else:
            new_name = "Corrected PowerPoint.pptx"

        return send_file(
            output_path,
            as_attachment=True,
            download_name=new_name,
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )

    except Exception as e:
        print(f"ERROR: {repr(e)}", flush=True)
        return jsonify({"error": str(e)}), 500

    finally:
        if input_path and os.path.exists(input_path):
            os.remove(input_path)


if __name__ == "__main__":
    app.run()
