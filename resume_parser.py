import os
import re
import fitz  # PyMuPDF
import pandas as pd

# Define education and region keywords
education_keywords = ["BTech", "MTech", "Bachelor", "Master", "MBA", "B.Sc", "M.Sc", "BCA", "MCA", "PhD"]
region_keywords = ["Bahrain", "Manama", "Riffa", "Muharraq", "Isa Town", "Sitra"]

def extract_text_from_pdf(path):
    doc = fitz.open(path)
    text = "\n".join(page.get_text() for page in doc)
    doc.close()
    return text

def extract_name_from_filename(filename):
    name_part = os.path.splitext(filename)[0]
    name_part = re.sub(r'[_\-.]', ' ', name_part)
    name_part = re.sub(r'\b(cv|resume|jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec|\d{2,4})\b', '', name_part, flags=re.IGNORECASE)
    name_part = re.sub(r'\s+', ' ', name_part).strip()
    words = name_part.split()
    name_clean = " ".join(w.capitalize() for w in words[:3])  # Get up to 3 words
    return name_clean

def extract_info(text, filename):
    name = extract_name_from_filename(filename)

    phone_match = re.search(r"(\+973\s*\d{6,8}|\d{8})", text)
    phone = phone_match.group(1) if phone_match else ""

    email_match = re.search(r"[\w\.-]+@[\w\.-]+", text)
    email = email_match.group() if email_match else ""

    region = next((region for region in region_keywords if region.lower() in text.lower()), "")

    edu_match = [word for word in education_keywords if word.lower() in text.lower()]
    education = "; ".join(set(edu_match)) if edu_match else ""

    grad_years = re.findall(r"\b(19|20)\d{2}\b", text)
    graduation_year = ", ".join(sorted(set(grad_years))) if grad_years else ""

    return {
        'Name': name,
        'Phone': phone,
        'Email': email,
        'Region': region,
        'Education': education,
        'Graduation Year': graduation_year
    }

def process_resumes(folder_path):
    records = []
    for filename in os.listdir(folder_path):
        if filename.lower().endswith('.pdf'):
            full_path = os.path.join(folder_path, filename)
            text = extract_text_from_pdf(full_path)
            info = extract_info(text, filename)
            info['File'] = filename
            records.append(info)
    return pd.DataFrame(records)

# ---- USAGE ----
folder = "resumes"  # Make sure your resumes are inside this folder
df = process_resumes(folder)
df.to_excel("Extracted_Resume_Data.xlsx", index=False)
print("✅ Data saved to Extracted_Resume_Data.xlsx")
