import os
import pdfplumber
import re

# Set the path to your folder of PDFs
folder_path = "C:/Users/User/Desktop/x"

# Regex pattern for emails
email_pattern = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'

# Store extracted emails
emails = {}

# Loop through each PDF file in the folder
for filename in os.listdir(folder_path):
    if filename.endswith(".pdf"):
        full_path = os.path.join(folder_path, filename)
        try:
            with pdfplumber.open(full_path) as pdf:
                text = ""
                for page in pdf.pages:
                    text += page.extract_text() or ""
                found_emails = re.findall(email_pattern, text)
                if found_emails:
                    emails[filename] = list(set(found_emails))  # remove duplicates
        except Exception as e:
            print(f"Failed to process {filename}: {e}")

# Print all extracted emails
for file, email_list in emails.items():
    print(f"\n{file}:")
    for email in email_list:
        print(f"  {email}")

import pandas as pd

rows = []
for file, email_list in emails.items():
    for email in email_list:
        rows.append({"Filename": file, "Email": email})

df = pd.DataFrame(rows)
df.to_excel("extracted_emails.xlsx", index=False)