import streamlit as st
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.print_page_options import PrintOptions
import base64
import time
import os
import zipfile
import io
import math
from pypdf import PdfWriter

# --- 1. PAGE SETUP & SIDEBAR ---
st.set_page_config(page_title="Bulk Challan Downloader", page_icon="🧾", layout="centered")

with st.sidebar:
    st.header("👨‍💻 About the Creator")
    st.markdown(
        """
        **Hasan Imam-Torongo**  
        ✉️ hasanimamofficial2912@gmail.com
        """
    )
    st.markdown("---")
    st.header("🔒 Privacy Note")
    st.info("All processing is done temporarily on the server. Your files and challan numbers are instantly deleted after your ZIP file is downloaded. No data is stored.")

# --- 2. MAIN HEADER & INSTRUCTIONS ---
st.title("🧾 Government Challan Bulk Downloader")
st.write("Upload your Excel file to automatically fetch, merge, and download multiple challans at once.")

# Expandable Instructions
with st.expander("📖 How to use this tool"):
    st.write("""
    1. Prepare an Excel file with a column named exactly **Challan Number**.
    2. Upload the file below.
    3. Click **Start Downloading** and wait for the process to finish.
    4. Download your ZIP file, which contains all individual PDFs plus a Master combined PDF!
    """)

# Button to download a perfect sample Excel file
def generate_sample_excel():
    output = io.BytesIO()
    sample_df = pd.DataFrame({"Challan Number": ["1234567890", "0987654321"]})
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        sample_df.to_excel(writer, index=False, sheet_name='Sheet1')
    return output.getvalue()

st.download_button(
    label="📄 Download Sample Excel Template",
    data=generate_sample_excel(),
    file_name="sample_challan_template.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.markdown("---")

# --- 3. FILE UPLOAD & DATA CLEANING ---
uploaded_file = st.file_uploader("Upload your Excel file (.xlsx)", type=['xlsx'])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    
    if 'Challan Number' not in df.columns:
        st.error("⚠️ Your Excel file MUST have a column header named exactly 'Challan Number'. Please download the sample template above.")
    else:
        # Clean the data: Remove empty rows and strip accidental spaces
        df = df.dropna(subset=['Challan Number'])
        df['Challan Number'] = df['Challan Number'].astype(str).str.strip()
        total_challans = len(df)
        
        # Calculate Estimated Time (approx 8 seconds per challan)
        estimated_minutes = math.ceil((total_challans * 8) / 60)
        
        st.success(f"✅ Found {total_challans} valid challans ready to download!")
        st.info(f"⏱️ **Estimated processing time:** ~{estimated_minutes} minute(s). Please do not refresh the page once started.")
        
        # --- 4. THE DOWNLOAD ENGINE ---
        if st.button("🚀 Start Downloading", use_container_width=True):
            
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            if not os.path.exists("temp_challans"):
                os.makedirs("temp_challans")
                
            # --- UPDATED: Server-Safe Chrome Setup with Anti-Lag ---
            options = webdriver.ChromeOptions()
            options.add_argument('--headless') 
            options.add_argument('--no-sandbox')
            options.add_argument('--disable-dev-shm-usage')
            options.add_argument('--disable-gpu')
            options.add_argument('--window-size=1920,1080') 
            
            # Make the server look like a normal Windows computer to bypass basic blocks
            options.add_argument('user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36')
            
            driver = webdriver.Chrome(options=options)
            
            # Force it to fail after 30 seconds instead of hanging forever
            driver.set_page_load_timeout(30)
            # -------------------------------------------------------
            
            pdf_file_paths = [] 
            successful_challans = []
            failed_challans = []
            
            for index, row in df.iterrows():
                challan_no = str(row['Challan Number'])
                status_text.text(f"Fetching Challan: {challan_no} ({index + 1} of {total_challans})...")
                
                try:
                    direct_url = f"https://challanverification.finance.gov.bd/echalan/details.php?challanNo={challan_no}"
                    driver.get(direct_url)
                    
                    time.sleep(5) # Give the government site time to load the data
                    
                    # Hide scrollbars for a clean PDF
                    driver.execute_script("document.documentElement.style.overflow = 'hidden';")
                    driver.execute_script("document.body.style.overflow = 'hidden';") 
                    
                    print_options = PrintOptions()
                    print_options.scale = 0.85 
                    print_options.background = True 
                    
                    pdf_data = driver.print_page(print_options)
                    
                    file_name = f"temp_challans/Challan_{challan_no}.pdf"
                    with open(file_name, "wb") as file:
                        file.write(base64.b64decode(pdf_data))
                        
                    pdf_file_paths.append(file_name)
                    successful_challans.append(challan_no)
                    
                except Exception as e:
                    # If it takes longer than 30s or gets blocked, it skips to here safely
                    failed_challans.append(challan_no)
                
                progress_bar.progress((index + 1) / total_challans)

            driver.quit()
            
            # --- 5. REPORTING & PACKAGING ---
            if len(successful_challans) > 0:
                status_text.text("📄 Merging all successful challans into a Master PDF...")
                merger = PdfWriter()
                
                for file_path in pdf_file_paths:
                    merger.append(file_path)
                    
                combined_pdf_path = "temp_challans/00_All_Challans_Combined.pdf"
                merger.write(combined_pdf_path)
                merger.close()
                pdf_file_paths.append(combined_pdf_path)

                status_text.text("📦 Packaging everything into a ZIP file...")
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w") as zip_file:
                    for file_path in pdf_file_paths:
                        file_name_only = os.path.basename(file_path)
                        zip_file.write(file_path, file_name_only)
                        
                status_text.empty()
                
                # Show Final Report
                st.markdown("### 📊 Final Report")
                st.success(f"✅ Successfully downloaded: **{len(successful_challans)}**")
                
                if failed_challans:
                    st.error(f"❌ Failed to download (Timeout/Blocked): **{len(failed_challans)}**")
                    with st.expander("View failed challan numbers"):
                        st.write(", ".join(failed_challans))
                
                st.download_button(
                    label="📥 Download ZIP Folder",
                    data=zip_buffer.getvalue(),
                    file_name="All_Challans_With_Master_Copy.zip",
                    mime="application/zip",
                    use_container_width=True
                )
                
                # Cleanup
                for file_path in pdf_file_paths:
                    if os.path.exists(file_path):
                        os.remove(file_path)
                os.rmdir("temp_challans")
            else:
                status_text.empty()
                st.error("❌ Failed to download any challans. The server might be blocking the connection or taking too long.")
