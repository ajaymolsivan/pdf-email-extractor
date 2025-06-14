import os
import zipfile
import re
import pdfplumber
import pandas as pd
from flask import Flask, render_template, request, send_file

app = Flask(__name__)

UPLOAD_FOLDER = os.path.join(os.getcwd(), "uploads")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

email_pattern = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        uploaded_file = request.files.get("zipfile")
        if uploaded_file and uploaded_file.filename.endswith(".zip"):
            zip_path = os.path.join(UPLOAD_FOLDER, uploaded_file.filename)
            uploaded_file.save(zip_path)

            # Unzip contents
            extract_folder = os.path.join(UPLOAD_FOLDER, "pdfs")
            os.makedirs(extract_folder, exist_ok=True)
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(extract_folder)

            # Extract emails
            emails = []
            for fname in os.listdir(extract_folder):
                if fname.endswith(".pdf"):
                    fpath = os.path.join(extract_folder, fname)
                    try:
                        with pdfplumber.open(fpath) as pdf:
                            text = ''.join([p.extract_text() or '' for p in pdf.pages])
                            found = list(set(re.findall(email_pattern, text)))
                            for e in found:
                                emails.append({"File": fname, "Email": e})
                    except Exception as e:
                        print(f"Error with {fname}: {e}")

            # Save to Excel
            
            output_path = os.path.join(UPLOAD_FOLDER, "extracted_emails.xlsx")
            df = pd.DataFrame(emails)
            df.to_excel(output_path, index=False)
            return send_file(output_path, as_attachment=True)
        

    return render_template("index.html")

if __name__ == "__main__":
    app.run(debug=True)
