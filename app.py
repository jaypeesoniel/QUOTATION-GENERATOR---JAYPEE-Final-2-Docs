import os
import tempfile
import shutil
from datetime import datetime
from flask import Flask, render_template, request, send_file
from docx import Document
from docx.shared import Pt

app = Flask(__name__)

# ====== AUTO-DETECT BASE DIRECTORY ======
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
GENERATED_FOLDER = os.path.join(BASE_DIR, "generated_docs")
TEMPLATES_FOLDER = os.path.join(BASE_DIR, "quotation_templates")
# ========================================

# --- Function to replace placeholders ---
def replace_placeholder(doc, placeholder, new_text, font_name="Cambria", font_size=12, bold=False):
    for p in doc.paragraphs:
        if placeholder in p.text:
            for run in p.runs:
                if placeholder in run.text:
                    run.text = run.text.replace(placeholder, new_text)
                    run.font.name = font_name
                    run.font.size = Pt(font_size)
                    run.bold = bold
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                replace_placeholder(cell, placeholder, new_text, font_name, font_size, bold)

@app.route('/', methods=['GET', 'POST'])
def index():
    # Ensure folders exist (Render-safe)
    os.makedirs(GENERATED_FOLDER, exist_ok=True)
    os.makedirs(TEMPLATES_FOLDER, exist_ok=True)

    # Auto-detect available templates
    models = [f.replace(".docx", "") for f in os.listdir(TEMPLATES_FOLDER) if f.endswith(".docx")]

    if request.method == 'POST':
        model = request.form.get('model_name').strip()
        client = request.form.get('client_name').strip().upper()
        address = request.form.get('address').strip().upper()
        contact = request.form.get('contact_person').strip().upper()
        current_date = datetime.now().strftime("%B %d, %Y")

        template_path = os.path.join(TEMPLATES_FOLDER, f"{model}.docx")
        if not os.path.exists(template_path):
            return f"❌ Walang template file para sa model '{model}'.<br><a href='/'>Balik</a>"

        # Format address (multi-line)
        address_lines = address.splitlines()
        formatted_address = "\n".join(line.strip() for line in address_lines if line.strip())

        # Temporary files
        temp_dir = tempfile.mkdtemp()
        temp_docx = os.path.join(temp_
