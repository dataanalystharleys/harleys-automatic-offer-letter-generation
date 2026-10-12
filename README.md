# Harley's Automatic Offer Letter Generation System

An automated HR offer letter generation web portal built for **Harley's India Pvt. Ltd.**

This system takes candidate details (either entered via a web form or bulk-uploaded through a standardized Excel sheet) and dynamically generates official offer letters in **Word (.docx)** and **PDF (.pdf)** formats while strictly preserving all corporate branding, logos, header/footer elements, and styling.

---

## 🌟 Key Features

1. **Exact Letterhead & Branding Preservation:**
   - Retains 100% of all company logos, header vector media, margins, and official footer credentials (`GSTIN: 36AAFCH0754C1ZY`, address, care@harleys.in).
2. **Dynamic Placeholders / Changeable Fields:**
   - Candidate Name, Salutation, Address, Mobile, Email
   - Designation / Role, Annual Compensation (CTC)
   - **Auto-formatting for Indian Currency:** E.g., `5000000` $\rightarrow$ `Rs. 50,00,000/-`
   - **Automatic Salary in Words:** E.g., `Fifty Lakh only/-`
   - Date of Joining, Offer Issue Date, Reporting Location & Time
   - HR Signatory Authority
3. **Bulk Generation via Excel:**
   - Download pre-styled `Harley_Candidate_Template.xlsx` template.
   - Upload filled Excel or CSV sheet with hundreds of candidate records.
   - Live editable data table preview before generation.
   - Generates all offer letters into a structured, downloadable ZIP archive (`DOCX/` and `PDF/` folders).
4. **Dual Output:**
   - Instant `.docx` and `.pdf` generation.

---

## 📋 Changeable Fields Mapping

| Field Name | Variable in Template | Original Sample Value | Excel Column Name |
| :--- | :--- | :--- | :--- |
| **Candidate Name** | `{{ candidate_name }}` | `Abhishek Rajput` | `Candidate Name` |
| **Salutation** | `{{ salutation }}` | `Mr.` | `Salutation` |
| **Address** | `{{ address }}` | `Marine Drive Patna, Bihar 843323` | `Address` |
| **Phone** | `{{ phone }}` | `+91 9267778873` | `Phone` |
| **Email** | `{{ email }}` | `abhi@gmail.com` | `Email` |
| **Designation** | `{{ designation }}` | `COO` | `Designation` |
| **Company Name** | `{{ company_name }}` | `HARLEY INDIA PVT LTD` | `Company Name` |
| **Annual Salary (CTC)** | `{{ salary_amount }}` | `50,00,000` | `Annual Salary (INR)` |
| **Salary in Words** | `{{ salary_in_words }}` | `Fifty Lakh only/-` | `Salary in Words` |
| **Date of Joining** | `{{ joining_date }}` | `20- 10- 2026` | `Date of Joining` |
| **Reporting Location** | `{{ reporting_location }}` | `Nanakramguda` | `Reporting Location` |
| **Reporting Time** | `{{ reporting_time }}` | `10.00 AM` | `Reporting Time` |
| **Offer Date** | `{{ offer_date }}` | `24-Sep-2026` | `Offer Date` |
| **HR Signatory** | `{{ hr_signatory }}` | `Sruti Associate - HR` | `HR Signatory` |

---

## 🚀 Quick Start (Local Windows)

### 1. One-Click Launch
Double-click `start.bat` in this folder. It will activate the virtual environment and open your web browser at `http://localhost:8501`.

### 2. Manual Command Line
```powershell
cd "C:\Users\harle\.gemini\antigravity\scratch\harleys-automatic-offer-letter-generation"
.\.venv\Scripts\activate
streamlit run app.py
```

---

## 🌐 How to Deploy Publicly for HR Team

### Option A: Streamlit Community Cloud (Recommended - 100% Free)
1. Push this project folder to a GitHub repository (public or private).
2. Go to [share.streamlit.io](https://share.streamlit.io) and sign in with GitHub.
3. Click **"New app"**, select your repository, set main file to `app.py`, and click **Deploy**.
4. Streamlit Community Cloud automatically reads `requirements.txt` and `packages.txt` (which installs `libreoffice` for PDF conversion).
5. You will get a permanent public HTTPS link (e.g., `https://harleys-offer-letters.streamlit.app`) to share with HR!

### Option B: Render / Railway / Cloud Run (Docker)
The included `Dockerfile` builds a production Linux container with Python 3.11 and LibreOffice pre-configured:
```bash
docker build -t harleys-offer-letter-app .
docker run -p 8501:8501 harleys-offer-letter-app
```

### Option C: Instant Temporary Public Link (Local Tunnel)
To immediately share your running local instance with HR for review:
```powershell
# Using cloudflared or localtunnel
npx localtunnel --port 8501
```
This generates an immediate public HTTPS URL to test with the HR team right away.

---

## 📁 Project Structure

```
harleys-automatic-offer-letter-generation/
│
├── app.py                      # Main Streamlit web application
├── generator.py                # Core docx/pdf generation & currency calculation engine
├── template_offer_letter.docx  # Master Word template with Jinja tags (all logos preserved)
├── sample_candidates.xlsx      # Standardized sample Excel template for HR
├── requirements.txt            # Python dependencies
├── packages.txt                # Linux apt dependencies for cloud PDF conversion
├── Dockerfile                  # Container definition for public cloud deployment
├── start.bat                   # 1-click Windows launcher
└── README.md                   # System documentation
```
