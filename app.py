import os
from flask import Flask, render_template, request, send_file
from docx import Document
from docx.shared import Pt
from datetime import datetime
import tempfile
import pypandoc

app = Flask(__name__)

# ====== PATHS ======
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATES_FOLDER = os.path.join(BASE_DIR, "quotation_templates")
os.makedirs(TEMPLATES_FOLDER, exist_ok=True)
# ===================

def replace_placeholder(doc, placeholder, new_text, font_name="Cambria", font_size=12, bold=False):
    """Palitan ang placeholder sa buong document (kasama tables)."""
    for p in doc.paragraphs:
        if placeholder in p.text:
            for run in p.runs:
                if placeholder in run.text:
                    run.text = run.text.replace(placeholder, new_text)
                    run.font.name = font_name
                    run.font.size = Pt(font_size)
                    run.bold = bold
    # Tables support
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                replace_placeholder(cell, placeholder, new_text, font_name, font_size, bold)

@app.route("/", methods=["GET", "POST"])
def index():
    # List of available templates (.docx)
    models = [f.replace(".docx", "") for f in os.listdir(TEMPLATES_FOLDER) if f.endswith(".docx")]

    if request.method == "POST":
        model = request.form.get("model_name", "").strip()
        client = request.form.get("client_name", "").strip().upper()
        address = request.form.get("address", "").strip().upper()
        contact = request.form.get("contact_person", "").strip().upper()
        date_value = request.form.get("date", "").strip()

        # Kung walang date input, gumamit ng current date
        if not date_value:
            date_value = datetime.now().strftime("%B %d, %Y")

        # Hanapin ang tamang template
        template_path = os.path.join(TEMPLATES_FOLDER, f"{model}.docx")
        if not os.path.exists(template_path):
            return f"❌ Walang template file para sa model '{model}'.<br><a href='/'>Balik</a>"

        # Load DOCX at palitan placeholders
        doc = Document(template_path)
        replace_placeholder(doc, "{{CLIENT_NAME}}", client, bold=True)
        replace_placeholder(doc, "{{ADDRESS}}", address)
        replace_placeholder(doc, "{{CONTACT_PERSON}}", contact, bold=True)
        replace_placeholder(doc, "{{DATE}}", date_value)
        replace_placeholder(doc, "ATTENTION:", "ATTENTION:", bold=True)

        # Gawa ng temp file at convert sa PDF
        with tempfile.TemporaryDirectory() as tmpdir:
            temp_docx = os.path.join(tmpdir, f"{client}_quotation.docx")
            temp_pdf = os.path.join(tmpdir, f"{client}_quotation.pdf")
            doc.save(temp_docx)

            try:
                pypandoc.convert_file(temp_docx, "pdf", outputfile=temp_pdf, extra_args=["--standalone"])
            except Exception as e:
                return f"⚠️ PDF conversion failed: {e}"

            return send_file(temp_pdf, as_attachment=True, download_name=f"{client}_quotation.pdf")

    return render_template("forms.html", models=models)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
