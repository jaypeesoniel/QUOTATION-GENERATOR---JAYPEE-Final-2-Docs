import os
from flask import Flask, render_template, request, send_file
from docx import Document
import tempfile

app = Flask(__name__)

# Folder kung saan nakalagay ang mga quotation templates
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATES_FOLDER = os.path.join(BASE_DIR, "quotation_templates")

# Ensure na existing folder kahit sa server
os.makedirs(TEMPLATES_FOLDER, exist_ok=True)

@app.route("/")
def index():
    # List ng available .docx templates
    models = []
    if os.path.exists(TEMPLATES_FOLDER):
        models = [f.replace(".docx", "") for f in os.listdir(TEMPLATES_FOLDER) if f.endswith(".docx")]
    # ⚠️ Change to your actual HTML name (forms.html)
    return render_template("index.html", models=models)

@app.route("/generate", methods=["POST"])
def generate():
    client_name = request.form.get("client_name", "")
    address = request.form.get("address", "")
    contact_person = request.form.get("contact_person", "")
    model = request.form.get("model", "")
    price = request.form.get("price", "")
    validity = request.form.get("validity", "")
    promo = request.form.get("promo", "")

    template_path = os.path.join(TEMPLATES_FOLDER, f"{model}.docx")

    if not os.path.exists(template_path):
        return f"Template file not found for model: {model}", 404

    # Load template
    doc = Document(template_path)

    # Replace placeholders
    replacements = {
        "{{client_name}}": client_name,
        "{{address}}": address,
        "{{contact_person}}": contact_person,
        "{{model}}": model,
        "{{price}}": price,
        "{{validity}}": validity,
        "{{promo}}": promo
    }

    for p in doc.paragraphs:
        for key, val in replacements.items():
            if key in p.text:
                p.text = p.text.replace(key, val)

    # Save temporary files
    with tempfile.TemporaryDirectory() as tmpdir:
        temp_docx = os.path.join(tmpdir, f"{client_name}_quotation.docx")
        doc.save(temp_docx)

        # ✅ Try to convert to PDF, fallback to DOCX if fails
        try:
            import pypandoc
            temp_pdf = os.path.join(tmpdir, f"{client_name}_quotation.pdf")
            pypa
