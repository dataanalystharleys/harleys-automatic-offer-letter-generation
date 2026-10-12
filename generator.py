import os
import shutil
import zipfile
from datetime import datetime
import pandas as pd
from docxtpl import DocxTemplate

def indian_number_to_words(number):
    """Convert a numeric amount into words according to the Indian numbering system."""
    try:
        clean_num = str(number).replace(',', '').replace('₹', '').replace('Rs.', '').replace('/-', '').strip()
        n = int(round(float(clean_num)))
    except Exception:
        return str(number)

    if n == 0:
        return "Zero only/-"

    units = ["", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine",
             "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen",
             "Seventeen", "Eighteen", "Nineteen"]
    tens = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"]

    def two_digits(num):
        if num < 20:
            return units[num]
        return tens[num // 10] + (" " + units[num % 10] if num % 10 != 0 else "")

    def three_digits(num):
        hundred = num // 100
        rem = num % 100
        res = ""
        if hundred > 0:
            res += units[hundred] + " Hundred"
            if rem > 0:
                res += " and "
        if rem > 0:
            res += two_digits(rem)
        return res

    parts = []
    crore = n // 10000000
    rem = n % 10000000
    if crore > 0:
        parts.append(f"{three_digits(crore)} Crore")

    lakh = rem // 100000
    rem = rem % 100000
    if lakh > 0:
        parts.append(f"{two_digits(lakh)} Lakh")

    thousand = rem // 1000
    rem = rem % 1000
    if thousand > 0:
        parts.append(f"{two_digits(thousand)} Thousand")

    if rem > 0:
        parts.append(f"{three_digits(rem)}")

    words = " ".join(parts).strip()
    return f"{words} only/-"

def format_indian_currency(number):
    """Format a number into Indian comma separated currency format, e.g. 50,00,000."""
    try:
        clean_num = str(number).replace(',', '').replace('₹', '').replace('Rs.', '').replace('/-', '').strip()
        s = str(int(round(float(clean_num))))
    except Exception:
        return str(number)
    if len(s) <= 3:
        return s
    last3 = s[-3:]
    remaining = s[:-3]
    res = ""
    while len(remaining) > 2:
        res = "," + remaining[-2:] + res
        remaining = remaining[:-2]
    return remaining + res + "," + last3

def calculate_salary_breakdown(annual_salary, pt_monthly_val=200, epf_val=0):
    """
    Standard Harley's India Pvt Ltd salary structure calculation (Annexure C):
    - Basic: 50% of CTC
    - HRA: 20% of CTC
    - Special Allowance: 30% of CTC
    - Professional Tax: ₹200 / month (₹2,400 / year)
    - In-Hand Salary: CTC - Deductions
    """
    try:
        clean_num = str(annual_salary).replace(',', '').replace('₹', '').replace('Rs.', '').replace('/-', '').strip()
        ctc_y = int(round(float(clean_num)))
    except Exception:
        ctc_y = 1020000

    ctc_m = int(round(ctc_y / 12.0))

    basic_y = int(round(ctc_y * 0.50))
    basic_m = int(round(basic_y / 12.0))

    hra_y = int(round(ctc_y * 0.20))
    hra_m = int(round(hra_y / 12.0))

    special_y = ctc_y - basic_y - hra_y
    special_m = ctc_m - basic_m - hra_m

    gross_y = basic_y + hra_y + special_y
    gross_m = basic_m + hra_m + special_m

    pt_m = pt_monthly_val
    pt_y = pt_m * 12

    epf_m = epf_val
    epf_y = epf_m * 12

    emp_pf_m = epf_val
    emp_pf_y = emp_pf_m * 12

    total_ded_y = pt_y + epf_y
    total_ded_m = pt_m + epf_m

    in_hand_y = gross_y - total_ded_y
    in_hand_m = gross_m - total_ded_m

    return {
        "ctc_annual_num": ctc_y,
        "ctc_monthly_num": ctc_m,
        "ctc_monthly": format_indian_currency(ctc_m),
        "basic_yearly": format_indian_currency(basic_y),
        "basic_monthly": format_indian_currency(basic_m),
        "hra_yearly": format_indian_currency(hra_y),
        "hra_monthly": format_indian_currency(hra_m),
        "special_yearly": format_indian_currency(special_y),
        "special_monthly": format_indian_currency(special_m),
        "gross_yearly": format_indian_currency(gross_y),
        "gross_monthly": format_indian_currency(gross_m),
        "epf_yearly": format_indian_currency(epf_y) if epf_y > 0 else "0",
        "epf_monthly": format_indian_currency(epf_m) if epf_m > 0 else "0",
        "emp_pf_yearly": format_indian_currency(emp_pf_y) if emp_pf_y > 0 else "0",
        "emp_pf_monthly": format_indian_currency(emp_pf_m) if emp_pf_m > 0 else "0",
        "pt_yearly": format_indian_currency(pt_y),
        "pt_monthly": format_indian_currency(pt_m),
        "total_deductions_yearly": format_indian_currency(total_ded_y),
        "total_deductions_monthly": format_indian_currency(total_ded_m),
        "in_hand_yearly": format_indian_currency(in_hand_y),
        "in_hand_monthly": format_indian_currency(in_hand_m)
    }

def format_date_str(val, fmt="%d-%b-%Y"):
    if pd.isna(val) or not val:
        return datetime.today().strftime(fmt)
    if isinstance(val, (datetime, pd.Timestamp)):
        return val.strftime(fmt)
    return str(val).strip()

def generate_document(template_path, context, output_docx_path):
    """Render any docx template with Jinja context."""
    tpl = DocxTemplate(template_path)
    tpl.render(context)
    tpl.save(output_docx_path)
    return output_docx_path

# Alias for backwards compatibility
generate_offer_letter = generate_document
generate_appointment_letter = generate_document

def convert_docx_to_pdf(docx_path, output_pdf_path):
    """
    Convert DOCX to PDF.
    Uses Word COM automation on Windows if available, or headless libreoffice if on Linux.
    Returns (True, None) on success or (False, error_message).
    """
    abs_docx = os.path.abspath(docx_path)
    abs_pdf = os.path.abspath(output_pdf_path)

    # 1. Try Word COM on Windows
    try:
        import win32com.client as win32
        word = win32.gencache.EnsureDispatch('Word.Application')
        word.Visible = False
        try:
            doc = word.Documents.Open(abs_docx)
            # wdFormatPDF = 17
            doc.SaveAs(abs_pdf, FileFormat=17)
            doc.Close(False)
            return True, None
        finally:
            try:
                word.Quit()
            except Exception:
                pass
    except Exception as e_word:
        # 2. Try headless LibreOffice if available
        import subprocess
        try:
            cmd = ["soffice", "--headless", "--convert-to", "pdf", "--outdir", os.path.dirname(abs_pdf), abs_docx]
            res = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=45)
            if res.returncode == 0 and os.path.exists(abs_pdf):
                return True, None
        except Exception:
            pass
        return False, f"PDF engine error: {str(e_word)}"

def sanitize_filename(name):
    return "".join(c for c in name if c.isalnum() or c in (' ', '-', '_')).strip()
