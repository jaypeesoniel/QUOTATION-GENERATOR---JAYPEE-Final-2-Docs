from flask import Flask, render_template, request, send_file
from PyPDF2 import PdfReader, PdfWriter
from datetime import datetime
import os, io, traceback

app = Flask(__name__)

# ====== PATH NG PDF TEMPLATE MO ======
PDF_TEMPLATE = os.path.join("quotation_templates", "ECOSYS M2540dn - 2.pdf")
# ====================================

@app.context_processor
def inject_year():
    return {'year': datetime.now().year}


def fill_pdf(template_path, data):
    reader = PdfReader(template_path)
    writer = PdfWriter()

    for page in reader.pages:
        writer.add_page(page)

    # Only update fields if form fields exist
    try:
        if reader.get_fields():
            writer.update_page_form_field_values(writer.pages[0], data)
    except Exception as e:
        print("⚠️ PDF field update error:", e)

    # Flatten (make fields read-only)
    for j in range(len(writer.pages)):
        page_obj = writer.pages[j]
        if "/Annots" in page_obj:
            for annot in page_obj["/Annots"]:
                obj = annot.get_object()
                if "/T" in obj:
                    obj.update({"/Ff": 1})  # mark as read-only

    output = io.BytesIO()
    writer.write(output)
    output.seek(0)
    return output


@app.route("/", methods=["GET", "POST"])
def index():
    models = [
        f.replace(".pdf", "") for f in os.listdir("quotation_templates") if f.endswith(".pdf")
    ]

    if request.method == "POST":
        try:
            model = request.form.get("model_name").strip()
            client = request.form.get("client_name").strip().upper()
            address = request.form.get("address").strip().upper()
            contact = request.form.get("contact_person").strip().upper()
            current_date = datetime.now().strftime("%B %d, %Y")

            data = {
                "DATE": current_date,
                "CLIENT_NAME": client,
                "ADDRESS": address,
                "CONTACT_PERSON": contact,
            }

            template_path = os.path.join("quotation_templates", f"{model}.pdf")
            if not os.path.exists(template_path):
                return f"❌ Walang PDF template na '{model}'.<br><a href='/'>Balik</a>"

            filled_pdf = fill_pdf(template_path, data)
            filename = f"Quotation - {client}.pdf"
            return send_file(filled_pdf, as_attachment=True, download_name=filename)

        except Exception as e:
            error_trace = traceback.format_exc()
            print("⚠️ Internal error:\n", error_trace)
            return f"<h3>❌ Error habang ginagawa ang PDF:</h3><pre>{error_trace}</pre>"

    return render_template("form.html", models=models)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
