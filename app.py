from flask import Flask, request, send_file
from pptx import Presentation
import io

app = Flask(__name__)

@app.route("/")
def home():
    return "PowerPoint AI Editor is running!"

@app.route("/edit", methods=["POST"])
def edit_powerpoint():
    if "file" not in request.files:
        return {"error": "No PowerPoint file received"}, 400

    uploaded_file = request.files["file"]

    presentation = Presentation(uploaded_file)

    output = io.BytesIO()
    presentation.save(output)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="corrected-powerpoint.pptx",
        mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )
