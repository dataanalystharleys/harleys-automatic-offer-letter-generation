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

def format_date_str(val):
    if pd.isna(val) or not val:
        return datetime.today().strftime("%d-%b-%Y")
    if isinstance(val, (datetime, pd.Timestamp)):
        return val.strftime("%d-%b-%Y")
    return str(val).strip()

def generate_offer_letter(template_path, context, output_docx_path):
    """Render the Word template using the provided context dictionary."""
    tpl = DocxTemplate(template_path)
    tpl.render(context)
    tpl.save(output_docx_path)
    return output_docx_path

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
            res = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=30)
            if res.returncode == 0 and os.path.exists(abs_pdf):
                return True, None
        except Exception:
            pass
        return False, f"PDF engine error: {str(e_word)}"

def sanitize_filename(name):
    return "".join(c for c in name if c.isalnum() or c in (' ', '-', '_')).strip()
