import os
from flask import Flask, request, render_template, send_from_directory
from docx2pdf import convert
from pdf2docx import Converter
import pythoncom

app = Flask(__name__)

# Ensure an output directory exists for converted files
output_dir = os.path.join(os.getcwd(), 'uploads')
os.makedirs(output_dir, exist_ok=True)

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/convert_docx_to_pdf', methods=['POST'])
def convert_docx_to_pdf():
    docx_file = request.files.get('docx_file')
    if docx_file:
        docx_file_path = os.path.join(output_dir, docx_file.filename)
        docx_file.save(docx_file_path)

        pdf_file_path = os.path.splitext(docx_file_path)[0] + ".pdf"

        try:
            pythoncom.CoInitialize()
            convert(docx_file_path, pdf_file_path)
            return send_from_directory(output_dir, os.path.basename(pdf_file_path), as_attachment=True)
        except Exception as e:
            return f"Error during conversion: {str(e)}"
        finally:
            pythoncom.CoUninitialize()

@app.route('/convert_pdf_to_docx', methods=['POST'])
def convert_pdf_to_docx():
    pdf_file = request.files.get('pdf_file')
    if pdf_file:
        pdf_file_path = os.path.join(output_dir, pdf_file.filename)
        pdf_file.save(pdf_file_path)

        docx_file_path = os.path.splitext(pdf_file_path)[0] + ".docx"

        try:
            converter = Converter(pdf_file_path)
            converter.convert(docx_file_path)
            converter.close()
            return send_from_directory(output_dir, os.path.basename(docx_file_path), as_attachment=True)
        except Exception as e:
            return f"Error during conversion: {str(e)}"

if __name__ == '__main__':
    app.run(debug=True)
