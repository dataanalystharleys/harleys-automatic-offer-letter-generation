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
    page_title="Harley's HR Document Generator",
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
    .preview-card {
        background-color: #F8FAFC;
        border: 1px solid #E2E8F0;
        padding: 1rem;
        border-radius: 0.5rem;
        margin-top: 0.5rem;
    }
    .salary-table {
        width: 100%;
        font-size: 0.85rem;
        margin-top: 0.5rem;
    }
</style>
""", unsafe_allow_html=True)

# Application Paths
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_OFFER = os.path.join(BASE_DIR, "template_offer_letter.docx")
TEMPLATE_APPOINTMENT = os.path.join(BASE_DIR, "template_appointment_letter.docx")
SAMPLE_EXCEL = os.path.join(BASE_DIR, "sample_candidates.xlsx")

# Ensure templates exist
if not os.path.exists(TEMPLATE_OFFER):
    st.error(f"Offer Letter template missing at `{TEMPLATE_OFFER}`!")
if not os.path.exists(TEMPLATE_APPOINTMENT):
    st.error(f"Appointment Letter template missing at `{TEMPLATE_APPOINTMENT}`!")

# Sidebar
with st.sidebar:
    st.image("https://img.icons8.com/fluency/96/document--v1.png", width=64)
    st.title("Harley's HR Portal")
    st.caption("Offer & Appointment Letter Generation System")
    st.markdown("---")
    
    st.markdown("### 🏢 Company Profile")
    st.write("**Company:** Harley's India Pvt. Ltd")
    st.write("**Location:** Nanakramguda, Hyderabad")
    st.write("**Branch:** Mumbai - CPU / Pan India")
    st.write("**GSTIN:** 36AAFCH0754C1ZY")
    st.write("**Official Domain:** www.harleys.in")
    
    st.markdown("---")
    st.markdown("### 📥 Resources & Templates")
    if os.path.exists(SAMPLE_EXCEL):
        with open(SAMPLE_EXCEL, "rb") as f:
            st.download_button(
                label="📊 Download Excel Master Template",
                data=f.read(),
                file_name="Harley_Candidates_Master_Template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="Download standardized Excel template with sample candidate rows"
            )
            
    if os.path.exists(TEMPLATE_OFFER):
        with open(TEMPLATE_OFFER, "rb") as f:
            st.download_button(
                label="📝 Download Offer Letter Template (.docx)",
                data=f.read(),
                file_name="template_offer_letter.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

    if os.path.exists(TEMPLATE_APPOINTMENT):
        with open(TEMPLATE_APPOINTMENT, "rb") as f:
            st.download_button(
                label="📝 Download Appointment Letter Template (.docx)",
                data=f.read(),
                file_name="template_appointment_letter.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

# Header
st.markdown('<div class="main-header">Harley\'s Automatic HR Document Generator</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Generate official <b>Offer Letters</b> and <b>Appointment Letters</b> in DOCX and PDF with automatic Indian currency calculations, Annexure C salary breakdowns, and bulk Excel batch processing.</div>', unsafe_allow_html=True)

# Main Navigation Tabs
tab1, tab2, tab3 = st.tabs([
    "👤 Single Candidate Generator",
    "📁 Bulk Batch Generation (Excel)",
    "ℹ️ Document & Field Reference"
])

# =========================================================================
# TAB 1: Single Candidate Generator
# =========================================================================
with tab1:
    st.subheader("Generate Candidate HR Documents")
    
    doc_type_choice = st.radio(
        "Select Document(s) to Generate:",
        ["Both (Offer Letter + Appointment Letter)", "Offer Letter Only", "Appointment Letter Only"],
        horizontal=True
    )
    
    st.markdown("---")
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("##### 👤 Candidate Profile")
        salutation = st.selectbox("Salutation", ["Mr.", "Ms.", "Mrs.", "Dr."], index=0)
        candidate_name = st.text_input("Candidate Full Name*", value="Roshan Machindra Mokal")
        email = st.text_input("Email Address*", value="roshan0mahendra@gmail.com")
        phone = st.text_input("Mobile / Contact Number*", value="+91 9022229535")
        address = st.text_area(
            "Candidate Full Address*",
            value="1007, tower 5, cluster 5, Blossom City, Near Rayan International School, Runwal Garden City, Usarghar, Dombivli 421204",
            height=90
        )
        
    with col2:
        st.markdown("##### 💼 Job & Compensation")
        designation = st.text_input("Designation / Role*", value="Chief Production Master")
        company_name = st.text_input("Company Name*", value="Harley’s India Private Limited")
        
        salary_numeric = st.number_input(
            "Annual Cost to Company (CTC in INR)*",
            min_value=50000,
            max_value=1000000000,
            value=1020000,
            step=50000,
            help="Enter annual compensation. Formatting, words, and Annexure C breakdown are auto-calculated."
        )
        
        auto_curr = generator.format_indian_currency(salary_numeric)
        auto_words = generator.indian_number_to_words(salary_numeric)
        breakdown = generator.calculate_salary_breakdown(salary_numeric)
        
        st.markdown(f"""
        <div class="preview-card">
            <span class="badge">Auto-Calculated</span><br>
            <b>Annual CTC:</b> Rs. {auto_curr}/- ({auto_words})<br>
            <b>Monthly Gross:</b> ₹{breakdown['gross_monthly']} | <b>Net In-Hand:</b> ₹{breakdown['in_hand_monthly']}/month<br>
            <details style="margin-top:0.4rem; font-size:0.85rem;">
                <summary><b>View Annexure C Structure Breakdown</b></summary>
                • Basic (50%): ₹{breakdown['basic_yearly']} / yr (₹{breakdown['basic_monthly']}/mo)<br>
                • HRA (20%): ₹{breakdown['hra_yearly']} / yr (₹{breakdown['hra_monthly']}/mo)<br>
                • Special Allowance (30%): ₹{breakdown['special_yearly']} / yr (₹{breakdown['special_monthly']}/mo)<br>
                • Professional Tax (PT): ₹{breakdown['pt_yearly']} / yr (₹{breakdown['pt_monthly']}/mo)
            </details>
        </div>
        """, unsafe_allow_html=True)
        
        salary_words_custom = st.text_input(
            "Salary in Words (Override if needed)",
            value=auto_words
        )
        
    st.markdown("---")
    col3, col4, col5 = st.columns(3)
    
    with col3:
        st.markdown("##### 📅 Dates & Probation")
        issue_date = st.date_input("Letter Issue Date*", value=datetime.today())
        joining_date = st.date_input("Date of Joining (w.e.f.)*", value=datetime.today())
        probation_period = st.text_input("Probation Period*", value="Three (3) months")
        
    with col4:
        st.markdown("##### 📍 Locations")
        work_location = st.text_input("Work Location*", value="Mumbai - CPU")
        reporting_location = st.text_input("Reporting Location*", value="Mumbai")
        reporting_time = st.text_input("Reporting Time*", value="10:00 a.m.")
        
    with col5:
        st.markdown("##### ✍️ Authority Signatory")
        signatory_name = st.text_input("Signatory Name*", value="K Harishanker Goud")
        signatory_title = st.text_input("Signatory Designation*", value="Authorized Signatory")
        
    st.markdown("---")
    
    # Generate action
    gen_btn = st.button("🚀 Generate Document(s)", type="primary", use_container_width=True)
    
    if gen_btn:
        if not candidate_name.strip():
            st.error("Please enter candidate name!")
        else:
            with st.spinner("Generating document(s)..."):
                clean_name = generator.sanitize_filename(candidate_name)
                formatted_issue_date = issue_date.strftime("%d-%b-%Y")
                formatted_joining_date = joining_date.strftime("%d-%b-%Y")
                
                # Context dictionary for template rendering
                context = {
                    "candidate_name": candidate_name.strip(),
                    "salutation": salutation,
                    "address": address.strip(),
                    "phone": phone.strip(),
                    "email": email.strip(),
                    "designation": designation.strip(),
                    "company_name": company_name.strip(),
                    "salary_amount": auto_curr,
                    "salary_in_words": salary_words_custom.strip(),
                    "offer_date": formatted_issue_date,
                    "appointment_date": formatted_issue_date,
                    "joining_date": formatted_joining_date,
                    "work_location": work_location.strip(),
                    "reporting_location": reporting_location.strip(),
                    "reporting_time": reporting_time.strip(),
                    "probation_period": probation_period.strip(),
                    "signatory_name": signatory_name.strip(),
                    "signatory_title": signatory_title.strip(),
                    "hr_signatory": f"{signatory_name.strip()} {signatory_title.strip()}",
                    **breakdown
                }
                
                temp_dir = tempfile.mkdtemp()
                generated_files = []
                
                # 1. Generate Offer Letter if selected
                if "Offer Letter" in doc_type_choice or "Both" in doc_type_choice:
                    offer_docx_name = f"Offer_Letter_{clean_name}.docx"
                    offer_pdf_name = f"Offer_Letter_{clean_name}.pdf"
                    offer_docx_path = os.path.join(temp_dir, offer_docx_name)
                    offer_pdf_path = os.path.join(temp_dir, offer_pdf_name)
                    
                    generator.generate_document(TEMPLATE_OFFER, context, offer_docx_path)
                    pdf_ok, _ = generator.convert_docx_to_pdf(offer_docx_path, offer_pdf_path)
                    
                    generated_files.append(("Offer Letter", "DOCX", offer_docx_name, offer_docx_path))
                    if pdf_ok and os.path.exists(offer_pdf_path):
                        generated_files.append(("Offer Letter", "PDF", offer_pdf_name, offer_pdf_path))
                        
                # 2. Generate Appointment Letter if selected
                if "Appointment Letter" in doc_type_choice or "Both" in doc_type_choice:
                    appt_docx_name = f"Appointment_Letter_{clean_name}.docx"
                    appt_pdf_name = f"Appointment_Letter_{clean_name}.pdf"
                    appt_docx_path = os.path.join(temp_dir, appt_docx_name)
                    appt_pdf_path = os.path.join(temp_dir, appt_pdf_name)
                    
                    generator.generate_document(TEMPLATE_APPOINTMENT, context, appt_docx_path)
                    pdf_ok, _ = generator.convert_docx_to_pdf(appt_docx_path, appt_pdf_path)
                    
                    generated_files.append(("Appointment Letter", "DOCX", appt_docx_name, appt_docx_path))
                    if pdf_ok and os.path.exists(appt_pdf_path):
                        generated_files.append(("Appointment Letter", "PDF", appt_pdf_name, appt_pdf_path))
                        
                st.success(f"🎉 Documents generated successfully for **{candidate_name}**!")
                
                # Present downloads
                st.markdown("##### 📥 Download Generated Files:")
                d_cols = st.columns(len(generated_files) if generated_files else 1)
                for idx, (doc_cat, doc_fmt, filename, filepath) in enumerate(generated_files):
                    with d_cols[idx]:
                        with open(filepath, "rb") as f_out:
                            mime_type = "application/pdf" if doc_fmt == "PDF" else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            st.download_button(
                                label=f"📥 {doc_cat} ({doc_fmt})",
                                data=f_out.read(),
                                file_name=filename,
                                mime=mime_type,
                                use_container_width=True
                            )
                            
                # If multiple files generated, provide 1-click ZIP bundle
                if len(generated_files) > 1:
                    zip_buf = io.BytesIO()
                    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as z:
                        for _, _, fn, fp in generated_files:
                            z.write(fp, arcname=fn)
                    zip_buf.seek(0)
                    st.download_button(
                        label=f"📦 Download Both as ZIP Archive ({clean_name}_Documents.zip)",
                        data=zip_buf.getvalue(),
                        file_name=f"{clean_name}_HR_Documents.zip",
                        mime="application/zip",
                        type="primary",
                        use_container_width=True
                    )


# =========================================================================
# TAB 2: Bulk Generation from Excel Sheet
# =========================================================================
with tab2:
    st.subheader("Bulk Generation from Excel Sheet")
    st.markdown("""
    Upload your Excel/CSV sheet containing candidate records. The system will generate **Offer Letters**, **Appointment Letters**, or **Both**, and package them into a neatly structured ZIP archive!
    """)
    
    col_up1, col_up2 = st.columns([2, 1])
    with col_up1:
        uploaded_file = st.file_uploader(
            "Upload Candidate Excel / CSV File",
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
                    file_name="Harley_Candidates_Master_Template.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

    if uploaded_file is not None:
        try:
            if uploaded_file.name.endswith(".csv"):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)
                
            st.markdown(f"##### 📋 Preview Data ({len(df)} candidate rows detected)")
            edited_df = st.data_editor(df, use_container_width=True, num_rows="dynamic")
            
            st.markdown("##### ⚙️ Generation Settings")
            g_col1, g_col2, g_col3 = st.columns(3)
            with g_col1:
                batch_doc_type = st.selectbox(
                    "Documents to Generate",
                    ["Both (Offer + Appointment)", "Offer Letters Only", "Appointment Letters Only"]
                )
            with g_col2:
                include_docx = st.checkbox("Include DOCX files", value=True)
            with g_col3:
                include_pdf = st.checkbox("Include PDF files", value=True)
                
            if st.button("⚡ Generate All Documents (ZIP Archive)", type="primary"):
                if len(edited_df) == 0:
                    st.error("No candidate rows found!")
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
                            status_text.text(f"Processing candidate {idx+1}/{total}: {c_name}...")
                            
                            salut = str(row.get("Salutation", "Mr.")).strip()
                            addr = str(row.get("Address", "")).strip()
                            ph = str(row.get("Phone", "")).strip()
                            em = str(row.get("Email", "")).strip()
                            desig = str(row.get("Designation", "Team Member")).strip()
                            comp = str(row.get("Company Name", "Harley’s India Private Limited")).strip()
                            
                            raw_salary = row.get("Annual Salary (INR)", 0)
                            sal_amt = generator.format_indian_currency(raw_salary)
                            
                            raw_words = row.get("Salary in Words", "")
                            if pd.isna(raw_words) or not str(raw_words).strip():
                                sal_words = generator.indian_number_to_words(raw_salary)
                            else:
                                sal_words = str(raw_words).strip()
                                
                            breakdown = generator.calculate_salary_breakdown(raw_salary)
                            
                            w_loc = str(row.get("Work Location", "Mumbai - CPU")).strip()
                            r_loc = str(row.get("Reporting Location", "Mumbai")).strip()
                            r_time = str(row.get("Reporting Time", "10:00 a.m.")).strip()
                            prob = str(row.get("Probation Period", "Three (3) months")).strip()
                            
                            join_d = generator.format_date_str(row.get("Date of Joining", datetime.today()))
                            iss_d = generator.format_date_str(row.get("Issue Date", datetime.today()))
                            sig_name = str(row.get("Signatory Name", "K Harishanker Goud")).strip()
                            sig_title = str(row.get("Signatory Title", "Authorized Signatory")).strip()
                            
                            ctx = {
                                "candidate_name": c_name,
                                "salutation": salut,
                                "address": addr,
                                "phone": ph,
                                "email": em,
                                "designation": desig,
                                "company_name": comp,
                                "salary_amount": sal_amt,
                                "salary_in_words": sal_words,
                                "offer_date": iss_d,
                                "appointment_date": iss_d,
                                "joining_date": join_d,
                                "work_location": w_loc,
                                "reporting_location": r_loc,
                                "reporting_time": r_time,
                                "probation_period": prob,
                                "signatory_name": sig_name,
                                "signatory_title": sig_title,
                                "hr_signatory": f"{sig_name} {sig_title}",
                                **breakdown
                            }
                            
                            # Row-level override if specified in sheet
                            row_doc_pref = str(row.get("Document Type", "Both")).strip().lower()
                            
                            gen_offer = "offer" in batch_doc_type.lower() or "both" in batch_doc_type.lower()
                            gen_appt = "appointment" in batch_doc_type.lower() or "both" in batch_doc_type.lower()
                            
                            if "offer" in row_doc_pref and "appointment" not in row_doc_pref and "both" not in row_doc_pref:
                                gen_appt = False
                            elif "appointment" in row_doc_pref and "offer" not in row_doc_pref and "both" not in row_doc_pref:
                                gen_offer = False
                                
                            if gen_offer:
                                o_docx_fn = f"Offer_Letter_{clean_name}_{idx+1}.docx"
                                o_docx_fp = os.path.join(temp_dir, o_docx_fn)
                                generator.generate_document(TEMPLATE_OFFER, ctx, o_docx_fp)
                                if include_docx:
                                    zip_file.write(o_docx_fp, arcname=f"Offer_Letters/DOCX/{o_docx_fn}")
                                if include_pdf:
                                    o_pdf_fn = f"Offer_Letter_{clean_name}_{idx+1}.pdf"
                                    o_pdf_fp = os.path.join(temp_dir, o_pdf_fn)
                                    ok, _ = generator.convert_docx_to_pdf(o_docx_fp, o_pdf_fp)
                                    if ok and os.path.exists(o_pdf_fp):
                                        zip_file.write(o_pdf_fp, arcname=f"Offer_Letters/PDF/{o_pdf_fn}")
                                        
                            if gen_appt:
                                a_docx_fn = f"Appointment_Letter_{clean_name}_{idx+1}.docx"
                                a_docx_fp = os.path.join(temp_dir, a_docx_fn)
                                generator.generate_document(TEMPLATE_APPOINTMENT, ctx, a_docx_fp)
                                if include_docx:
                                    zip_file.write(a_docx_fp, arcname=f"Appointment_Letters/DOCX/{a_docx_fn}")
                                if include_pdf:
                                    a_pdf_fn = f"Appointment_Letter_{clean_name}_{idx+1}.pdf"
                                    a_pdf_fp = os.path.join(temp_dir, a_pdf_fn)
                                    ok, _ = generator.convert_docx_to_pdf(a_docx_fp, a_pdf_fp)
                                    if ok and os.path.exists(a_pdf_fp):
                                        zip_file.write(a_pdf_fp, arcname=f"Appointment_Letters/PDF/{a_pdf_fn}")
                                        
                            success_count += 1
                            progress_bar.progress((idx + 1) / total)
                            
                    status_text.text("✨ Generation completed!")
                    st.success(f"Successfully processed {success_count} candidate records!")
                    
                    zip_buffer.seek(0)
                    now_str = datetime.now().strftime("%Y%m%d_%H%M%S")
                    st.download_button(
                        label="📦 Download All Generated Documents (ZIP Archive)",
                        data=zip_buffer.getvalue(),
                        file_name=f"Harley_HR_Documents_{now_str}.zip",
                        mime="application/zip",
                        type="primary",
                        use_container_width=True
                    )
        except Exception as e:
            st.error(f"Error reading file: {e}")


# =========================================================================
# TAB 3: Document & Field Reference
# =========================================================================
with tab3:
    st.subheader("Document Comparison & Template Specifications")
    st.markdown("""
    Both documents use your official Harley's India Pvt Ltd letterhead styling, logos, and registered office footer.
    """)
    
    col_ref1, col_ref2 = st.columns(2)
    with col_ref1:
        st.markdown("#### 📄 1. Offer Letter")
        st.info("""
        - **Purpose**: Initial conditional employment proposal extended to candidate.
        - **Length**: 1 page concise official proposal.
        - **Key Clauses**:
          - Annual CTC and statutory deductions mention
          - Reporting location & time
          - First-day document submission checklist (AADHAR, PAN, Relieving letter, Payslips, Bank details, Photos)
          - HR sign-off
        """)
        
    with col_ref2:
        st.markdown("#### 📜 2. Appointment Letter")
        st.info("""
        - **Purpose**: Comprehensive, legally binding employment agreement issued on/after joining.
        - **Length**: 4-5 pages with full contractual terms.
        - **Key Clauses & Annexures**:
          - **Annexure A**: 19 contractual clauses (Probation, Deputation, Confidentiality, Non-Compete, Notice Period, Arbitration).
          - **Annexure B**: Roles & Responsibilities reference.
          - **Annexure C**: Detailed Salary Structure Table (Basic 50%, HRA 20%, Special 30%, PT deduction, Net In-Hand).
        """)

    st.markdown("---")
    st.markdown("##### 📌 Master Field Mapping Matrix")
    matrix_data = [
        {"Field Name": "Candidate Name", "Offer Letter": "Yes", "Appointment Letter": "Yes", "Excel Column": "Candidate Name"},
        {"Field Name": "Designation", "Offer Letter": "Yes", "Appointment Letter": "Yes", "Excel Column": "Designation"},
        {"Field Name": "Annual CTC (INR)", "Offer Letter": "Yes", "Appointment Letter": "Yes", "Excel Column": "Annual Salary (INR)"},
        {"Field Name": "Salary in Words", "Offer Letter": "Yes", "Appointment Letter": "Yes", "Excel Column": "Salary in Words"},
        {"Field Name": "Date of Joining", "Offer Letter": "Yes", "Appointment Letter": "Yes", "Excel Column": "Date of Joining"},
        {"Field Name": "Letter Issue Date", "Offer Letter": "Yes", "Appointment Letter": "Yes", "Excel Column": "Issue Date"},
        {"Field Name": "Candidate Address", "Offer Letter": "Yes", "Appointment Letter": "Yes", "Excel Column": "Address"},
        {"Field Name": "Candidate Phone", "Offer Letter": "Yes", "Appointment Letter": "Yes", "Excel Column": "Phone"},
        {"Field Name": "Candidate Email", "Offer Letter": "Yes", "Appointment Letter": "Yes", "Excel Column": "Email"},
        {"Field Name": "Work Location", "Offer Letter": "N/A", "Appointment Letter": "Yes", "Excel Column": "Work Location"},
        {"Field Name": "Reporting Location", "Offer Letter": "Yes", "Appointment Letter": "Yes", "Excel Column": "Reporting Location"},
        {"Field Name": "Probation Period", "Offer Letter": "N/A", "Appointment Letter": "Yes", "Excel Column": "Probation Period"},
        {"Field Name": "Annexure C Table", "Offer Letter": "N/A", "Appointment Letter": "Yes (Auto)", "Excel Column": "Auto-calculated"},
        {"Field Name": "Authorized Signatory", "Offer Letter": "Yes", "Appointment Letter": "Yes", "Excel Column": "Signatory Name / Title"},
    ]
    st.table(pd.DataFrame(matrix_data))
