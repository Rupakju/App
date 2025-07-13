import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import datetime
import io
import zipfile
import os
import tempfile

# Page configuration
st.set_page_config(
    page_title="Invitation Letter Generator",
    page_icon="📝",
    layout="wide"
)

# Custom CSS for better styling
st.markdown("""
<style>
.main-header {
    text-align: center;
    color: #2c3e50;
    font-size: 2.5rem;
    margin-bottom: 2rem;
}
.sub-header {
    color: #34495e;
    font-size: 1.2rem;
    margin-bottom: 1rem;
}
.success-message {
    background-color: #d4edda;
    color: #155724;
    padding: 1rem;
    border-radius: 0.5rem;
    margin: 1rem 0;
}
</style>
""", unsafe_allow_html=True)

# Title
st.markdown('<h1 class="main-header">📝 Invitation Letter Generator</h1>', unsafe_allow_html=True)
st.markdown('<p class="sub-header">Generate professional invitation letters for visa applications</p>', unsafe_allow_html=True)

# Initialize session state
if 'generated_files' not in st.session_state:
    st.session_state.generated_files = []

def read_word_data(uploaded_file):
    """Read data from uploaded Word file"""
    try:
        doc = Document(uploaded_file)
        info = {}
        
        # Read from first table
        if doc.tables:
            table = doc.tables[0]
            for row in table.rows:
                if len(row.cells) < 2:
                    continue
                key = row.cells[0].text.strip()
                value = row.cells[1].text.strip()
                info[key] = value
        
        return info
    except Exception as e:
        st.error(f"Error reading file: {str(e)}")
        return None

def create_invitation_letter(data, header_image=None, footer_image=None, signature_image=None):
    """Create invitation letter document"""
    try:
        # Create a new Word document
        doc = Document()
        
        # Adjust margins
        section = doc.sections[0]
        section.left_margin = Inches(0.9)
        section.right_margin = Inches(0.8)
        section.header_distance = Inches(0.2)
        section.footer_distance = Inches(0.1)
        
        # Add header image
        if header_image:
            header = section.header
            header_paragraph = header.paragraphs[0]
            header_paragraph.alignment = 2  # Right-align
            header_run = header_paragraph.add_run()
            header_run.add_picture(header_image, width=Inches(3.44))
        
        # Add footer image
        if footer_image:
            footer = section.footer
            footer_paragraph = footer.paragraphs[0]
            footer_run = footer_paragraph.add_run()
            footer_run.add_picture(footer_image, width=Inches(6.75))
        
        # Set default font
        style = doc.styles['Normal']
        font = style.font
        font.name = 'Lato'
        font.size = Pt(11)
        
        # Add content
        current_date = datetime.now().strftime("%B %d, %Y")
        doc.add_paragraph(f"Date: {current_date}\n")
        
        # To paragraph
        to_para = doc.add_paragraph(f"To\n{data.get('Location of the Bangladesh Embassy that you are applying to (fill address)', 'N/A')}\n")
        
        # Subject line
        subject_para = doc.add_paragraph()
        subject_para.add_run("Subject: Request for business visa to ").bold = False
        subject_para.add_run(f"{data.get('Full Name \n(As it appears on passport)', 'N/A')}, ").bold = True
        subject_para.add_run("Passport No: ").bold = False
        subject_para.add_run(f"{data.get('Passport number', 'N/A')}, ").bold = True
        subject_para.add_run("Nationality: ").bold = False
        subject_para.add_run(f"{data.get('Nationality', 'N/A')}.\n").bold = True
        
        # Body content
        doc.add_paragraph("Dear Sir/Madam,").alignment = 3
        
        doc.add_paragraph("""Save the Children is an international development organization working in 120 countries around the world. The headquarters of Save the Children is located at St Vincent House, 30 Orange Street, London WC2H 7HH, United Kingdom. Save the Children is registered with the NGO Affairs Bureau in Bangladesh (registered number 2630, dated March 20, 2011).""").alignment = 3
        
        # Main body with bold formatting
        body_para = doc.add_paragraph()
        body_para.add_run(f"{data.get('Full Name \n(As it appears on passport)', 'N/A')}, ").bold = True
        body_para.add_run(f"{data.get('Job Title', 'N/A')}, of Save the Children has been invited to the Save the Children International office in Bangladesh to participate in meetings, training, and program activities from ")
        body_para.add_run(f"{data.get('Arrival Date in Bangladesh', 'N/A')} ").bold = True
        body_para.add_run("to ")
        body_para.add_run(f"{data.get('Departure Date', 'N/A')}. ").bold = True
        body_para.add_run("The Save the Children Bangladesh Country Office will ensure all logistical support.").bold = False
        body_para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        
        # Assistance paragraph
        assistance_para = doc.add_paragraph()
        assistance_para.add_run(f"Your kind assistance in granting a visa for {data.get('Full Name \n(As it appears on passport)', 'N/A')} ")
        assistance_para.add_run("to visit Bangladesh would be highly appreciated. Please contact at Cell no:  +8801913918618 and mail: ")
        assistance_para.add_run("sumon.paul@savethechildren.org ").bold = True
        assistance_para.add_run("if there is any query regarding the processing of this visa.")
        assistance_para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        
        doc.add_paragraph("Thank you for your kind assistance.").alignment = 3
        
        # Add signature image
        if signature_image:
            signature_paragraph = doc.add_paragraph()
            signature_run = signature_paragraph.add_run()
            signature_run.add_picture(signature_image, width=Inches(2.5))
        
        doc.add_paragraph(f"Sumon kumar Paul\nCoordinator - Administration\n")
        
        return doc
        
    except Exception as e:
        st.error(f"Error creating invitation letter: {str(e)}")
        return None

# Main interface
col1, col2 = st.columns([1, 1])

with col1:
    st.markdown("### 📁 Upload Files")
    
    # File uploads
    uploaded_files = st.file_uploader(
        "Upload Word files with applicant data",
        type=['docx'],
        accept_multiple_files=True,
        help="Select one or more Word files containing applicant information"
    )
    
    # Image uploads
    st.markdown("### 🖼️ Upload Images (Optional)")
    
    header_image = st.file_uploader(
        "Header Image",
        type=['png', 'jpg', 'jpeg'],
        help="Image for document header"
    )
    
    footer_image = st.file_uploader(
        "Footer Image", 
        type=['png', 'jpg', 'jpeg'],
        help="Image for document footer"
    )
    
    signature_image = st.file_uploader(
        "Signature Image",
        type=['png', 'jpg', 'jpeg'], 
        help="Signature image for the document"
    )

with col2:
    st.markdown("### ⚙️ Generation Options")
    
    # Generate button
    if st.button("🚀 Generate Invitation Letters", type="primary", use_container_width=True):
        if not uploaded_files:
            st.error("Please upload at least one Word file")
        else:
            # Progress bar
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            generated_files = []
            
            for i, uploaded_file in enumerate(uploaded_files):
                # Update progress
                progress = (i + 1) / len(uploaded_files)
                progress_bar.progress(progress)
                status_text.text(f"Processing {uploaded_file.name}...")
                
                # Read data
                data = read_word_data(uploaded_file)
                if data:
                    # Create invitation letter
                    doc = create_invitation_letter(
                        data, 
                        header_image=header_image, 
                        footer_image=footer_image,
                        signature_image=signature_image
                    )
                    
                    if doc:
                        # Save to memory
                        doc_buffer = io.BytesIO()
                        doc.save(doc_buffer)
                        doc_buffer.seek(0)
                        
                        # Generate filename
                        base_name = uploaded_file.name.replace('.docx', '')
                        filename = f"{base_name}_invitation.docx"
                        
                        generated_files.append({
                            'filename': filename,
                            'data': doc_buffer.getvalue(),
                            'applicant': data.get('Full Name \n(As it appears on passport)', 'Unknown')
                        })
            
            # Store in session state
            st.session_state.generated_files = generated_files
            
            # Clear progress
            progress_bar.empty()
            status_text.empty()
            
            # Success message
            if generated_files:
                st.success(f"✅ Successfully generated {len(generated_files)} invitation letters!")
            else:
                st.error("❌ No files could be processed")

# Download section
if st.session_state.generated_files:
    st.markdown("### 📥 Download Generated Files")
    
    # Create download options
    col1, col2 = st.columns([1, 1])
    
    with col1:
        # Individual downloads
        st.markdown("**Individual Downloads:**")
        for file_info in st.session_state.generated_files:
            st.download_button(
                label=f"📄 {file_info['applicant']}",
                data=file_info['data'],
                file_name=file_info['filename'],
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
    
    with col2:
        # Bulk download as ZIP
        st.markdown("**Bulk Download:**")
        
        # Create ZIP file
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            for file_info in st.session_state.generated_files:
                zip_file.writestr(file_info['filename'], file_info['data'])
        
        zip_buffer.seek(0)
        
        st.download_button(
            label="📦 Download All as ZIP",
            data=zip_buffer.getvalue(),
            file_name=f"invitation_letters_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
            mime="application/zip"
        )

# Instructions
with st.expander("📋 Instructions"):
    st.markdown("""
    ### How to use this tool:
    
    1. **Prepare your Word files**: Each file should contain a table with applicant information including:
       - Full Name (As it appears on passport)
       - Passport number
       - Nationality
       - Job Title
       - Arrival Date in Bangladesh
       - Departure Date
       - Location of the Bangladesh Embassy
    
    2. **Upload files**: Select one or more Word files containing applicant data
    
    3. **Upload images** (optional): Add header, footer, and signature images
    
    4. **Generate**: Click the generate button to create invitation letters
    
    5. **Download**: Download individual files or all files as a ZIP archive
    
    ### Supported formats:
    - Input: Word documents (.docx)
    - Images: PNG, JPG, JPEG
    - Output: Word documents (.docx)
    """)

# Footer
st.markdown("---")
st.markdown(
    '<p style="text-align: center; color: #7f8c8d;">Save the Children - Invitation Letter Generator</p>',
    unsafe_allow_html=True
)