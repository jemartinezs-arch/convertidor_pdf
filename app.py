from flask import Flask, request, send_file, render_template, jsonify
from services.pdf_service import PDFService
import os

app = Flask(__name__)

OUTPUT_FOLDER = "static/outputs"
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# ---------------- Rutas HTML ----------------

@app.route("/")
def home():
    return render_template("index.html")

@app.route("/pdf-a-word")
def pdf_word():
    return render_template("pdf-a-word.html")

@app.route("/pdf-a-png")
def pdf_png():
    return render_template("pdf-a-png.html")

@app.route("/privacy")
def privacy():
    return render_template("privacy.html")

@app.route("/terms")
def terms():
    return render_template("terms.html")

@app.route("/contact")
def contact():
    return render_template("contact.html")

@app.route("/gracias")
def gracias():
    return render_template("gracias.html")

@app.route("/feedback", methods=["POST"])
def feedback():
    data = request.get_json()
    print("Feedback recibido:", data.get("feedback"))
    return jsonify({"status": "ok"})

# ---------------- PDF → PNG ----------------

@app.route("/convert/png", methods=["POST"])
def pdf_to_png():
    try:
        if "file" not in request.files:
            return "No se envió ningún archivo", 400

        pdf_file = request.files["file"]

        if not pdf_file.filename.lower().endswith(".pdf"):
            return "Formato inválido. Solo PDFs.", 400

        zip_buffer = PDFService.pdf_to_png(pdf_file.read())

        return send_file(
            zip_buffer,
            as_attachment=True,
            download_name="imagenes.zip",
            mimetype="application/zip"
        )

    except Exception as e:
        print("ERROR PNG:", e)
        return "Error procesando el PDF.", 500

# ---------------- PDF → WORD ----------------

@app.route("/convert/word", methods=["POST"])
def pdf_to_word():
    try:
        if "file" not in request.files:
            return "No se envió archivo", 400

        pdf_file = request.files["file"]

        if not pdf_file.filename.lower().endswith(".pdf"):
            return "Formato inválido. Solo PDFs.", 400

        buffer = PDFService.pdf_to_word(pdf_file.read())

        return send_file(
            buffer,
            as_attachment=True,
            download_name="convertido.docx",
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except Exception as e:
        print("ERROR WORD:", e)
        return "Error procesando PDF a Word.", 500

# ---------------- Producción ----------------

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)