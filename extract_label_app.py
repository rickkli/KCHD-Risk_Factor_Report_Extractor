import streamlit as st
import csv
import pandas as pd
from datetime import datetime
import io

# Set page configuration
st.set_page_config(page_title="Risk Factor Report Extractor", layout="wide")

st.title("Risk Factor Report Extractor")
st.write("Enter the Facility Numbers, upload your CSV files, and download the cleaned Excel reports.")

# --- UI ELEMENT 1: Text Box for Facility Numbers ---
st.subheader("1. Enter Facility Numbers")
raw_copied_text = st.text_area(
    "Paste your Facility Numbers here.", 
    height=150, 
)

# --- UI ELEMENT 2: File Uploader ---
st.subheader("2. Upload CSV Files")
uploaded_files = st.file_uploader("Upload CSV files here.", type="csv", accept_multiple_files=True)

# --- UI ELEMENT 3: Process Button ---
if st.button("Process Data", type="primary"):
    
    # Validation
    if not raw_copied_text.strip():
        st.error("⚠️ Please enter at least one Facility Number.")
    elif not uploaded_files:
        st.error("⚠️ Please upload at least one CSV file.")
    else:
        with st.spinner("Processing data..."):
            
            # Clean the ID list
            filter_list = [line.strip() for line in raw_copied_text.strip().split('\n') if line.strip()]

            # 2. CONFIGURATION
            target_items = ['6', '7', '9A', '9B', '10', '11', '12B', '13A', '13B', '13C', '14', '16', '17', '18A', '18B', '19A', '20A', '20B', '21A', '21B', '25', '29', '31', '37', '47']
            
            sub_items_16 = [f"16{chr(i)}" for i in range(ord('A'), ord('J'))] 
            sub_items_17 = [f"17{chr(i)}" for i in range(ord('A'), ord('E'))] 
            
            extract_items = set(target_items + sub_items_16 + sub_items_17)
            
            status_map = {'In': 'in', 'Out': 'out', 'N/O': 'no', 'N/A': 'na'}
            
            reports = []
            match_counts = {fac: 0 for fac in filter_list}

            # 3. PROCESSING
            for uploaded_file in uploaded_files:
                current_school = None
                current_id = None
                current_fac_num = None
                current_report_data = None

                # Decode the uploaded file byte stream to string lines for the csv.reader
                file_content = uploaded_file.getvalue().decode('utf-8').splitlines()
                reader = csv.reader(file_content)
                
                for row in reader:
                    if not row: continue
                    
                    # Identify Facility Name
                    if row[0].strip() and len(row) > 2 and not row[1].strip() and not row[2].strip():
                        current_school = row[0].strip()
                        continue
                        
                    # Identify Inspection #, Facility #, and Date
                    if not row[0].strip() and len(row) >= 11 and row[1].strip() and not row[2].strip():
                        current_id = row[1].strip()
                        current_fac_num = row[7].strip()
                        current_date_str = row[10].strip()
                        
                        if current_fac_num in filter_list:
                            match_counts[current_fac_num] += 1
                            
                            try:
                                current_date = datetime.strptime(current_date_str, '%m/%d/%Y')
                            except ValueError:
                                current_date = pd.NaT
                                
                            current_report_data = {
                                'Facility Name': current_school,
                                'Facility Number': current_fac_num,
                                'Inspection #': current_id,
                                'Inspection Date': current_date_str,
                                '_parsed_date': current_date 
                            }
                            reports.append(current_report_data)
                        else:
                            current_report_data = None 
                        continue
                        
                    # Extract items
                    if current_report_data is not None and len(row) > 3 and row[2].strip():
                        item_num = row[2].strip()
                        item_num_clean = item_num.rstrip('.')
                        
                        if item_num in extract_items or item_num_clean in extract_items:
                            matched_item = item_num_clean if item_num_clean in extract_items else item_num
                            raw_status = next((val.strip() for val in reversed(row) if val.strip()), None)
                            
                            if raw_status in status_map:
                                current_report_data[matched_item] = status_map[raw_status]

            # 4. ROLL-UP LOGIC FOR 16 AND 17
            for rep in reports:
                stat_16 = [rep.get(sub) for sub in sub_items_16 if rep.get(sub) is not None]
                stat_17 = [rep.get(sub) for sub in sub_items_17 if rep.get(sub) is not None]
                
                if "out" in stat_16: rep['16'] = "out"
                elif "in" in stat_16: rep['16'] = "in"
                elif "no" in stat_16: rep['16'] = "no"
                elif "na" in stat_16: rep['16'] = "na"
                
                if "out" in stat_17: rep['17'] = "out"
                elif "in" in stat_17: rep['17'] = "in"
                elif "no" in stat_17: rep['17'] = "no"
                elif "na" in stat_17: rep['17'] = "na"

            # 5. FILTERING & DEDUPLICATION
            df = pd.DataFrame(reports)

            if not df.empty:
                df = df.sort_values(by=['Facility Number', '_parsed_date'], ascending=[True, False])
                df = df.drop_duplicates(subset=['Facility Number'], keep='first')
                df = df.drop(columns=['_parsed_date'])

            for item in target_items:
                if item not in df.columns: 
                    df[item] = ""

            cols = ['Facility Name', 'Facility Number', 'Inspection Date', 'Inspection #'] + target_items
            if not df.empty:
                df = df[cols]
            else:
                df = pd.DataFrame(columns=cols)

            # 6. GENERATE AUDIT REPORT
            audit_data = []
            for fac_num, count in match_counts.items():
                status = "OK"
                if count == 0:
                    status = "MISSING (Not found in CSV)"
                elif count > 1:
                    status = f"DUPLICATE ({count} reports found - Extracted Most Recent)"
                
                audit_data.append({'Facility Number': fac_num, 'Match Count': count, 'Audit Status': status})

            audit_df = pd.DataFrame(audit_data)
            
            # --- PREPARE FILES FOR DOWNLOAD ---
            # Helper function to convert dataframe to excel bytes in memory
            def convert_df_to_excel(dataframe):
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    dataframe.to_excel(writer, index=False, sheet_name='Sheet1')
                return output.getvalue()

            main_excel_bytes = convert_df_to_excel(df)
            audit_excel_bytes = convert_df_to_excel(audit_df)

            # Display completion metrics
            st.success("Extraction Complete!")
            col1, col2 = st.columns(2)
            col1.metric("Unique Facilities Extracted", len(df))
            col2.metric("Missing Facilities", len(audit_df[audit_df['Match Count'] == 0]))

            # --- UI ELEMENT 4: Download Buttons ---
            st.subheader("3. Download Results")
            st.write("Click the buttons below to download your processed Excel files.")
            
            # Stack download buttons vertically
            st.download_button(
                label="📥 Download Filtered Data",
                data=main_excel_bytes,
                file_name='Filtered_Inspections.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
            st.download_button(
                label="📥 Download Audit Log",
                data=audit_excel_bytes,
                file_name='Extraction_Audit_Log.xlsx',
                mime='application/vnd.openxmlformats-officedocument-spreadsheetml.sheet'
            )