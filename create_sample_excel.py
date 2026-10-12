import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

excel_path = "sample_candidates.xlsx"

data = [
    {
        "Candidate Name": "Roshan Machindra Mokal",
        "Salutation": "Mr.",
        "Designation": "Chief Production Master",
        "Annual Salary (INR)": 1020000,
        "Salary in Words": "Ten Lakh Twenty Thousand only/-",
        "Date of Joining": "05-Oct-2026",
        "Issue Date": "05-Oct-2026",
        "Document Type": "Both",
        "Address": "1007, tower 5, cluster 5, Blossom City, Near Rayan International School, Runwal Garden City, Usarghar, Dombivli 421204",
        "Phone": "+91 9022229535",
        "Email": "roshan0mahendra@gmail.com",
        "Work Location": "Mumbai - CPU",
        "Reporting Location": "Mumbai",
        "Reporting Time": "10:00 a.m.",
        "Probation Period": "Three (3) months",
        "Company Name": "Harley’s India Private Limited",
        "Signatory Name": "K Harishanker Goud",
        "Signatory Title": "Authorized Signatory"
    },
    {
        "Candidate Name": "Abhishek Rajput",
        "Salutation": "Mr.",
        "Designation": "COO",
        "Annual Salary (INR)": 5000000,
        "Salary in Words": "Fifty Lakh only/-",
        "Date of Joining": "20-Oct-2026",
        "Issue Date": "24-Sep-2026",
        "Document Type": "Both",
        "Address": "Marine Drive Patna, Bihar 843323",
        "Phone": "+91 9267778873",
        "Email": "abhi@gmail.com",
        "Work Location": "Hyderabad - Nanakramguda",
        "Reporting Location": "Nanakramguda",
        "Reporting Time": "10.00 AM",
        "Probation Period": "Three (3) months",
        "Company Name": "HARLEY INDIA PVT LTD",
        "Signatory Name": "Sruti Associate - HR",
        "Signatory Title": "Associate - HR"
    },
    {
        "Candidate Name": "Priya Sharma",
        "Salutation": "Ms.",
        "Designation": "Senior Full Stack Developer",
        "Annual Salary (INR)": 1800000,
        "Salary in Words": "Eighteen Lakh only/-",
        "Date of Joining": "01-Nov-2026",
        "Issue Date": "12-Oct-2026",
        "Document Type": "Both",
        "Address": "Plot 42, Hitech City, Hyderabad, Telangana 500081",
        "Phone": "+91 9876501234",
        "Email": "priya.sharma@example.com",
        "Work Location": "Hyderabad - Nanakramguda",
        "Reporting Location": "Nanakramguda",
        "Reporting Time": "10.00 AM",
        "Probation Period": "Three (3) months",
        "Company Name": "Harley’s India Private Limited",
        "Signatory Name": "K Harishanker Goud",
        "Signatory Title": "Authorized Signatory"
    }
]

df = pd.DataFrame(data)
df.to_excel(excel_path, index=False)

# Add styling to Excel file
wb = openpyxl.load_workbook(excel_path)
ws = wb.active

header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
border_thin = Border(
    left=Side(style='thin', color='DDDDDD'),
    right=Side(style='thin', color='DDDDDD'),
    top=Side(style='thin', color='DDDDDD'),
    bottom=Side(style='thin', color='DDDDDD')
)

for cell in ws[1]:
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

for row in ws.iter_rows(min_row=2):
    for cell in row:
        cell.font = Font(name="Calibri", size=10)
        cell.border = border_thin
        if cell.column_letter in ['D']: # Salary
            cell.number_format = '#,##,##0'

for col in ws.columns:
    max_len = max(len(str(cell.value or '')) for cell in col)
    col_letter = col[0].column_letter
    ws.column_dimensions[col_letter].width = max(min(max_len + 4, 35), 14)

wb.save(excel_path)
print("Updated Excel template created and styled successfully!")
