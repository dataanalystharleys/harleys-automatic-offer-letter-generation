import os
import io
import zipfile
import tempfile
from datetime import datetime
import pandas as pd
import streamlit as st

import generator

# Page configuration
st.set_page_config(
    page_title="Harley's Automatic Offer Letter Generator",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for polished HR styling
st.markdown("""
<style>
    .main-header {
        font-size: 2.2rem;
        font-weight: 700;
        color: #1E3A8A;
        margin-bottom: 0.2rem;
    }
    .sub-header {
        font-size: 1.05rem;
        color: #4B5563;
        margin-bottom: 1.5rem;
    }
    .badge {
        display: inline-block;
        padding: 0.2rem 0.6rem;
        border-radius: 9999px;
        font-size: 0.8rem;
        font-weight: 600;
        background-color: #DBEAFE;
        color: #1E40AF;
        margin-right: 0.4rem;
    }
    .success-box {
        background-color: #ECFDF5;
        border-left: 4px solid #10B981;
        padding: 1rem;
        border-radius: 0.375rem;
        margin: 1rem 0;
    }
    .preview-card {
        background-color: #F9FAFB;
        border: 1px solid #E5E7EB;
        padding: 1rem;
        border-radius: 0.5rem;
        margin-top: 0.5rem;
    }
</style>
""", unsafe_allow_html=True)

# Application Paths
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_DOCX = os.path.join(BASE_DIR, "template_offer_letter.docx")
SAMPLE_EXCEL = os.path.join(BASE_DIR, "sample_candidates.xlsx")

# Ensure template exists
if not os.path.exists(TEMPLATE_DOCX):
    st.error(f"Template not found at `{TEMPLATE_DOCX}`! Please ensure `template_offer_letter.docx` is present.")
    st.stop()

# Sidebar
with st.sidebar:
    st.image("https://img.icons8.com/fluency/96/document--v1.png", width=64)
    st.title("Harley's HR Portal")
    st.caption("Automatic Offer Letter Generation System")
    st.markdown("---")
    
    st.markdown("### 🏢 Company Profile")
    st.write("**Company:** Harley's India Pvt. Ltd")
    st.write("**Location:** Nanakramguda, Hyderabad")
    st.write("**CIN / GSTIN:** 36AAFCH0754C1ZY")
    st.write("**Official Domain:** www.harleys.in")
    
    st.markdown("---")
    st.markdown("### 📥 Resources")
    if os.path.exists(SAMPLE_EXCEL):
        with open(SAMPLE_EXCEL, "rb") as f:
            st.download_button(
                label="📊 Download Excel Template",
                data=f.read(),
                file_name="Harley_Candidate_Template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="Download standardized Excel template with sample candidate rows"
            )
            
    if os.path.exists(TEMPLATE_DOCX):
        with open(TEMPLATE_DOCX, "rb") as f:
            st.download_button(
                label="📝 Download Base Template (.docx)",
                data=f.read(),
                file_name="template_offer_letter.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                help="Download the underlying Word template with placeholders"
            )

# Header
st.markdown('<div class="main-header">Harley\'s Automatic Offer Letter Generator</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Generate official offer letters in <b>DOCX</b> and <b>PDF</b> with automatic Indian currency formatting, salary in words, and batch Excel processing.</div>', unsafe_allow_html=True)

# Main Navigation Tabs
tab1, tab2, tab3 = st.tabs([
    "👤 Single Candidate Offer Letter",
    "📁 Bulk Generation from Excel Sheet",
    "ℹ️ Changeable Fields & Template Guide"
])

# =========================================================================
# TAB 1: Single Candidate Offer Letter
# =========================================================================
with tab1:
    st.subheader("Candidate & Offer Details")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("##### 👤 Candidate Profile")
        salutation = st.selectbox("Salutation", ["Mr.", "Ms.", "Mrs.", "Dr."], index=0)
        candidate_name = st.text_input("Candidate Full Name*", value="Abhishek Rajput")
        email = st.text_input("Email Address*", value="abhi@gmail.com")
        phone = st.text_input("Mobile / Contact Number*", value="+91 9267778873")
        address = st.text_area("Candidate Address*", value="Marine Drive Patna, Bihar 843323", height=90)
        
    with col2:
        st.markdown("##### 💼 Job & Compensation")
        designation = st.text_input("Designation / Role*", value="COO")
        company_name = st.text_input("Company Name*", value="HARLEY INDIA PVT LTD")
        
        salary_numeric = st.number_input(
            "Annual CTC / Salary (in INR)*",
            min_value=50000,
            max_value=1000000000,
            value=5000000,
            step=50000,
            help="Enter annual compensation. Formatting and words are auto-calculated."
        )
        
        auto_curr = generator.format_indian_currency(salary_numeric)
        auto_words = generator.indian_number_to_words(salary_numeric)
        
        st.markdown(f"""
        <div class="preview-card">
            <span class="badge">Auto-Calculated</span><br>
            <b>Formatted:</b> Rs. {auto_curr}/-<br>
            <b>In Words:</b> {auto_words}
        </div>
        """, unsafe_allow_html=True)
        
        salary_words_custom = st.text_input(
            "Salary in Words (Customize if needed)",
            value=auto_words,
            help="Default is generated automatically based on Indian numbering system."
        )
        
    st.markdown("---")
    col3, col4, col5 = st.columns(3)
    
    with col3:
        st.markdown("##### 📅 Dates")
        offer_date = st.date_input("Offer Letter Issue Date", value=datetime.today())
        joining_date = st.date_input("Date of Joining*", value=datetime.today())
        
    with col4:
        st.markdown("##### 📍 Reporting Information")
        reporting_location = st.text_input("Reporting Location*", value="Nanakramguda")
        reporting_time = st.text_input("Reporting Time*", value="10.00 AM")
        
    with col5:
        st.markdown("##### ✍️ Authority")
        hr_signatory = st.text_input("HR Signatory*", value="Sruti Associate - HR")
        
    st.markdown("---")
    
    # Generate action
    gen_col1, gen_col2 = st.columns([1, 2])
    with gen_col1:
        generate_btn = st.button("🚀 Generate Offer Letter", type="primary", use_container_width=True)
        
    if generate_btn:
        if not candidate_name.strip():
            st.error("Please enter candidate name!")
        else:
            with st.spinner("Generating offer letter..."):
                clean_name = generator.sanitize_filename(candidate_name)
                formatted_offer_date = offer_date.strftime("%d-%b-%Y")
                formatted_joining_date = joining_date.strftime("%d-%m-%Y")
                
                context = {
                    "offer_date": formatted_offer_date,
                    "salutation": salutation,
                    "candidate_name": candidate_name.strip(),
                    "address": address.strip(),
                    "phone": phone.strip(),
                    "email": email.strip(),
                    "designation": designation.strip(),
                    "company_name": company_name.strip(),
                    "salary_amount": auto_curr,
                    "salary_in_words": salary_words_custom.strip(),
                    "reporting_location": reporting_location.strip(),
                    "reporting_time": reporting_time.strip(),
                    "joining_date": formatted_joining_date,
                    "hr_signatory": hr_signatory.strip()
                }
                
                temp_dir = tempfile.mkdtemp()
                docx_filename = f"Offer_Letter_{clean_name}.docx"
                pdf_filename = f"Offer_Letter_{clean_name}.pdf"
                docx_out = os.path.join(temp_dir, docx_filename)
                pdf_out = os.path.join(temp_dir, pdf_filename)
                
                # Generate DOCX
                generator.generate_offer_letter(TEMPLATE_DOCX, context, docx_out)
                
                # Attempt PDF conversion
                pdf_success, pdf_err = generator.convert_docx_to_pdf(docx_out, pdf_out)
                
                st.success(f"🎉 Offer letter generated successfully for **{candidate_name}**!")
                
                # Download buttons
                down_col1, down_col2 = st.columns(2)
                
                with down_col1:
                    with open(docx_out, "rb") as f_docx:
                        st.download_button(
                            label=f"📥 Download DOCX ({docx_filename})",
                            data=f_docx.read(),
                            file_name=docx_filename,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True
                        )
                        
                with down_col2:
                    if pdf_success and os.path.exists(pdf_out):
                        with open(pdf_out, "rb") as f_pdf:
                            st.download_button(
                                label=f"📥 Download PDF ({pdf_filename})",
                                data=f_pdf.read(),
                                file_name=pdf_filename,
                                mime="application/pdf",
                                use_container_width=True
                            )
                    else:
                        st.info("ℹ️ PDF direct conversion is available on Windows with Word / LibreOffice. DOCX file has been generated with all formatting and logos preserved.")


# =========================================================================
# TAB 2: Bulk Generation from Excel Sheet
# =========================================================================
with tab2:
    st.subheader("Bulk Generation from Excel Sheet")
    st.markdown("""
    HR can prepare an Excel sheet with multiple candidate details. Upload it here to generate all offer letters at once in a single downloadable ZIP archive!
    """)
    
    col_up1, col_up2 = st.columns([2, 1])
    with col_up1:
        uploaded_file = st.file_uploader(
            "Upload Filled Candidate Excel / CSV File",
            type=["xlsx", "xls", "csv"],
            help="Select the Excel or CSV file containing candidate rows."
        )
    with col_up2:
        st.markdown("<br>", unsafe_allow_html=True)
        if os.path.exists(SAMPLE_EXCEL):
            with open(SAMPLE_EXCEL, "rb") as f_ex:
                st.download_button(
                    label="📥 Download Template Excel File",
                    data=f_ex.read(),
                    file_name="Harley_Candidate_Template.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

    if uploaded_file is not None:
        try:
            if uploaded_file.name.endswith(".csv"):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)
                
            st.markdown(f"##### 📋 Preview Uploaded Data ({len(df)} candidate records found)")
            
            # Allow interactive in-table editing by HR
            edited_df = st.data_editor(df, use_container_width=True, num_rows="dynamic")
            
            st.markdown("##### ⚙️ Generation Options")
            gen_opts_col1, gen_opts_col2 = st.columns(2)
            with gen_opts_col1:
                include_docx = st.checkbox("Include DOCX files", value=True)
            with gen_opts_col2:
                include_pdf = st.checkbox("Include PDF files", value=True)
                
            if st.button("⚡ Generate All Offer Letters (ZIP Archive)", type="primary"):
                if len(edited_df) == 0:
                    st.error("No candidate rows found in the sheet!")
                else:
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    zip_buffer = io.BytesIO()
                    temp_dir = tempfile.mkdtemp()
                    
                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                        total = len(edited_df)
                        success_count = 0
                        
                        for idx, row in edited_df.iterrows():
                            c_name = str(row.get("Candidate Name", f"Candidate_{idx+1}")).strip()
                            clean_name = generator.sanitize_filename(c_name)
                            status_text.text(f"Processing ({idx+1}/{total}): {c_name}...")
                            
                            salut = str(row.get("Salutation", "Mr.")).strip()
                            addr = str(row.get("Address", "")).strip()
                            ph = str(row.get("Phone", "")).strip()
                            em = str(row.get("Email", "")).strip()
                            desig = str(row.get("Designation", "Team Member")).strip()
                            comp = str(row.get("Company Name", "HARLEY INDIA PVT LTD")).strip()
                            
                            raw_salary = row.get("Annual Salary (INR)", 0)
                            try:
                                sal_amt = generator.format_indian_currency(raw_salary)
                            except Exception:
                                sal_amt = str(raw_salary)
                                
                            raw_words = row.get("Salary in Words", "")
                            if pd.isna(raw_words) or not str(raw_words).strip():
                                sal_words = generator.indian_number_to_words(raw_salary)
                            else:
                                sal_words = str(raw_words).strip()
                                
                            loc = str(row.get("Reporting Location", "Nanakramguda")).strip()
                            rep_time = str(row.get("Reporting Time", "10.00 AM")).strip()
                            
                            join_d = generator.format_date_str(row.get("Date of Joining", datetime.today()))
                            off_d = generator.format_date_str(row.get("Offer Date", datetime.today()))
                            hr_sig = str(row.get("HR Signatory", "Sruti Associate - HR")).strip()
                            
                            ctx = {
                                "offer_date": off_d,
                                "salutation": salut,
                                "candidate_name": c_name,
                                "address": addr,
                                "phone": ph,
                                "email": em,
                                "designation": desig,
                                "company_name": comp,
                                "salary_amount": sal_amt,
                                "salary_in_words": sal_words,
                                "reporting_location": loc,
                                "reporting_time": rep_time,
                                "joining_date": join_d,
                                "hr_signatory": hr_sig
                            }
                            
                            docx_filename = f"Offer_Letter_{clean_name}_{idx+1}.docx"
                            docx_path = os.path.join(temp_dir, docx_filename)
                            generator.generate_offer_letter(TEMPLATE_DOCX, ctx, docx_path)
                            
                            if include_docx:
                                zip_file.write(docx_path, arcname=f"DOCX/{docx_filename}")
                                
                            if include_pdf:
                                pdf_filename = f"Offer_Letter_{clean_name}_{idx+1}.pdf"
                                pdf_path = os.path.join(temp_dir, pdf_filename)
                                ok_pdf, _ = generator.convert_docx_to_pdf(docx_path, pdf_path)
                                if ok_pdf and os.path.exists(pdf_path):
                                    zip_file.write(pdf_path, arcname=f"PDF/{pdf_filename}")
                                    
                            success_count += 1
                            progress_bar.progress((idx + 1) / total)
                            
                    status_text.text("✨ Generation completed!")
                    st.success(f"Successfully generated offer letters for {success_count} candidate(s)!")
                    
                    zip_buffer.seek(0)
                    now_str = datetime.now().strftime("%Y%m%d_%H%M%S")
                    st.download_button(
                        label="📦 Download All Offer Letters (ZIP Archive)",
                        data=zip_buffer.getvalue(),
                        file_name=f"Harley_Offer_Letters_{now_str}.zip",
                        mime="application/zip",
                        type="primary",
                        use_container_width=True
                    )
        except Exception as e:
            st.error(f"Error reading file: {e}")


# =========================================================================
# TAB 3: Changeable Fields & Template Guide
# =========================================================================
with tab3:
    st.subheader("Offer Letter Field Mapping & Template Specification")
    st.markdown("""
    The system reads your master template docx (`template_offer_letter.docx`) and injects the following fields without altering the layout, header logos, or footer information:
    """)
    
    mapping_data = [
        {"Field Name": "Candidate Name", "Jinja Placeholder": "{{ candidate_name }}", "Original Template Value": "Abhishek Rajput", "Excel Column": "Candidate Name"},
        {"Field Name": "Salutation", "Jinja Placeholder": "{{ salutation }}", "Original Template Value": "Mr.", "Excel Column": "Salutation"},
        {"Field Name": "Candidate Address", "Jinja Placeholder": "{{ address }}", "Original Template Value": "Marine Drive Patna, Bihar 843323", "Excel Column": "Address"},
        {"Field Name": "Candidate Phone", "Jinja Placeholder": "{{ phone }}", "Original Template Value": "+91 9267778873", "Excel Column": "Phone"},
        {"Field Name": "Candidate Email", "Jinja Placeholder": "{{ email }}", "Original Template Value": "abhi@gmail.com", "Excel Column": "Email"},
        {"Field Name": "Designation", "Jinja Placeholder": "{{ designation }}", "Original Template Value": "COO", "Excel Column": "Designation"},
        {"Field Name": "Company Name", "Jinja Placeholder": "{{ company_name }}", "Original Template Value": "HARLEY INDIA PVT LTD", "Excel Column": "Company Name"},
        {"Field Name": "Annual Salary (Amount)", "Jinja Placeholder": "{{ salary_amount }}", "Original Template Value": "50,00,000", "Excel Column": "Annual Salary (INR)"},
        {"Field Name": "Salary in Words", "Jinja Placeholder": "{{ salary_in_words }}", "Original Template Value": "Fifty Lakh only/-", "Excel Column": "Salary in Words"},
        {"Field Name": "Date of Joining", "Jinja Placeholder": "{{ joining_date }}", "Original Template Value": "20- 10- 2026", "Excel Column": "Date of Joining"},
        {"Field Name": "Reporting Location", "Jinja Placeholder": "{{ reporting_location }}", "Original Template Value": "Nanakramguda", "Excel Column": "Reporting Location"},
        {"Field Name": "Reporting Time", "Jinja Placeholder": "{{ reporting_time }}", "Original Template Value": "10.00 AM", "Excel Column": "Reporting Time"},
        {"Field Name": "Offer Issue Date", "Jinja Placeholder": "{{ offer_date }}", "Original Template Value": "24-Sep-2026", "Excel Column": "Offer Date"},
        {"Field Name": "HR Signatory", "Jinja Placeholder": "{{ hr_signatory }}", "Original Template Value": "Sruti Associate - HR", "Excel Column": "HR Signatory"},
    ]
    st.table(pd.DataFrame(mapping_data))
    
    st.markdown("##### 📌 Preserved Document Sections (Fixed Letterhead)")
    st.info("""
    - **Header & Company Logos**: The company logo, top branding banners, and vector assets are 100% preserved.
    - **Required Document Checklist**:
      - `> AADHAR and PAN cards`
      - `> Previous employment relieving letter, if applicable`
      - `> Payslips of last three months`
      - `> Bank Details`
      - `> Educational and experience certificates`
      - `> Photos - 2`
    - **Closing Wish**: `We wish you a great career with Harley’s`
    - **Official Footer**:
      - `Head Office: Harley’s India Pvt. Ltd, Survey No:55/E, Nanakramguda Road, Hyderabad, Telangana, India-500032. | GSTIN: 36AAFCH0754C1ZY`
      - `+91 80830 98888 | care@harleys.in | www.harleys.in`
    """)
