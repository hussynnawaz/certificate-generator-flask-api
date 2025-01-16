from flask import Flask, request, jsonify, send_file
from docx2pdf import convert
import os

app = Flask(__name__)

# Ensure the output folder exists
output_folder = "testing"
os.makedirs(output_folder, exist_ok=True)

@app.route("/wordtopdf", methods=["POST"])
def wordtopdf():
    try:
        # Check if a file is part of the request
        if 'file' not in request.files:
            return jsonify({"error": "No file provided"}), 400

        file = request.files['file']

        # Ensure the file has a .docx extension
        if not file.filename.endswith('.docx'):
            return jsonify({"error": "Only .docx files are supported"}), 400

        # Save the uploaded file temporarily
        input_path = os.path.join(output_folder, file.filename)
        file.save(input_path)

        # Define output PDF path
        output_pdf_path = os.path.splitext(input_path)[0] + ".pdf"

        # Convert .docx to .pdf
        convert(input_path, output_pdf_path)

        # Send the converted PDF as a response
        return send_file(output_pdf_path, as_attachment=True)

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(debug=True)
