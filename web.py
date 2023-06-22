from flask import Flask, request, send_file
from docx2pdf import convert
import os
from datetime import datetime
import pythoncom

app = Flask(__name__)

@app.route('/convert', methods=['POST'])
def convert_to_pdf():
    file = request.files['file']
    # Mendapatkan nama file yang dikirimkan oleh Laravel
    docx_filename = file.filename

    # Membuat nama file PDF yang unik berdasarkan tanggal dan waktu
    timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
    pdf_filename = f"output_{timestamp}.pdf"

    file.save(os.path.join(app.root_path, docx_filename))

    pythoncom.CoInitialize()  # Memanggil CoInitialize sebelum konversi

    convert(os.path.join(app.root_path, docx_filename), os.path.join(app.root_path, pdf_filename))

    pdf_path = os.path.join(app.root_path, pdf_filename)
    return send_file(pdf_path, as_attachment=True)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5001)
