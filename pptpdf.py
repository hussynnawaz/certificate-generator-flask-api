from flask import Flask, request, jsonify, send_file
import os
from comtypes.client import CreateObject
import comtypes

app = Flask(__name__)

# Ensure the output folder exists
output_folder = "testing"
os.makedirs(output_folder, exist_ok=True)

@app.route("/ppttopdf", methods=["POST"])
def ppttopdf():
    try:
        # Check if a file is part of the request
        if 'file' not in request.files:
            return jsonify({"error": "No file provided"}), 400

        file = request.files['file']

        # Ensure the file has a .ppt or .pptx extension
        if not (file.filename.endswith('.ppt') or file.filename.endswith('.pptx')):
            return jsonify({"error": "Only .ppt or .pptx files are supported"}), 400

        # Save the uploaded file temporarily
        input_path = os.path.abspath(os.path.join(output_folder, file.filename))
        file.save(input_path)

        # Define output PDF path
        output_pdf_path = os.path.splitext(input_path)[0] + ".pdf"

        # Initialize COM library
        comtypes.CoInitialize()

        # Convert .ppt/.pptx to .pdf using comtypes
        powerpoint = CreateObject("PowerPoint.Application")
        powerpoint.Visible = 1

        presentation = powerpoint.Presentations.Open(input_path, WithWindow=False)
        presentation.SaveAs(output_pdf_path, 32)  # 32 is the PDF format
        presentation.Close()
        powerpoint.Quit()

        # Send the converted PDF as a response
        return send_file(output_pdf_path, as_attachment=True)

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(debug=True)
