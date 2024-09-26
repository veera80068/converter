import os
from flask import Flask, request, render_template, send_from_directory
from docx2pdf import convert
from pdf2docx import Converter
import pythoncom
import logging

app = Flask(__name__)

# Set up logging
logging.basicConfig(level=logging.INFO)

# Use a different upload folder
UPLOAD_FOLDER = os.path.join("C:\\temp\\uploads")
OUTPUT_FOLDER = os.path.join("C:\\temp\\output")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/convert_docx_to_pdf', methods=['POST'])
def convert_docx_to_pdf():
    docx_file = request.files.get('docx_file')
    
    if not docx_file:
        return "No DOCX file uploaded.", 400
    
    # Save uploaded DOCX file to uploads folder
    docx_file_path = os.path.join(UPLOAD_FOLDER, docx_file.filename)
    
    try:
        docx_file.save(docx_file_path)  # Save uploaded file
        logging.info(f"DOCX file saved at: {docx_file_path}")
    except PermissionError:
        logging.error(f"Permission denied for file: {docx_file.filename}")
        return "Permission denied. Please check if the file is open or if you have permission to write in the uploads directory.", 403
    except Exception as e:
        logging.error(f"Error during saving the file: {str(e)}")
        return f"Error during saving the file: {str(e)}", 500

    # Generate output PDF file path
    pdf_file_path = os.path.join(OUTPUT_FOLDER, f"{os.path.splitext(docx_file.filename)[0]}.pdf")

    try:
        pythoncom.CoInitialize()  # Initialize COM for Windows
        # Convert DOCX to PDF
        convert(docx_file_path, pdf_file_path)
        logging.info(f"Conversion successful. PDF saved at: {pdf_file_path}")
        return send_from_directory(OUTPUT_FOLDER, os.path.basename(pdf_file_path), as_attachment=True)
    except Exception as e:
        logging.error(f"Error during conversion: {str(e)}")
        return f"Error during conversion: {str(e)}"
    finally:
        pythoncom.CoUninitialize()  # Uninitialize COM

@app.route('/convert_pdf_to_docx', methods=['POST'])
def convert_pdf_to_docx():
    pdf_file = request.files.get('pdf_file')
    
    if not pdf_file:
        return "No PDF file uploaded.", 400
    
    # Save uploaded PDF file to uploads folder
    pdf_file_path = os.path.join(UPLOAD_FOLDER, pdf_file.filename)
    
    try:
        pdf_file.save(pdf_file_path)
        logging.info(f"PDF file saved at: {pdf_file_path}")
    except Exception as e:
        logging.error(f"Error during saving the file: {str(e)}")
        return f"Error during saving the file: {str(e)}", 500

    # Generate output DOCX file path
    docx_file_path = os.path.join(OUTPUT_FOLDER, f"{os.path.splitext(pdf_file.filename)[0]}.docx")

    try:
        # Convert PDF to DOCX
        converter = Converter(pdf_file_path)
        converter.convert(docx_file_path, start=0, end=None)
        converter.close()
        logging.info(f"Conversion successful. DOCX saved at: {docx_file_path}")
        return send_from_directory(OUTPUT_FOLDER, os.path.basename(docx_file_path), as_attachment=True)
    except Exception as e:
        logging.error(f"Error during conversion: {str(e)}")
        return f"Error during conversion: {str(e)}"

if __name__ == '__main__':
    app.run(debug=True)
