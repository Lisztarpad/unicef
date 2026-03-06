import streamlit as st
import pandas as pd
from datetime import datetime
import io
import re
from openpyxl.utils import get_column_letter

# Set page configuration
st.set_page_config(page_title="Data Processing & Anomaly Detection", layout="wide")

st.title("Data Processing & Anomaly Detection App")
st.write("Please upload the required reports to find anomaly cases and their responsible persons.")

# File uploaders
col1, col2, col3 = st.columns(3)
with col1:
    file_intella = st.file_uploader("1. Intella Report (.csv)", type=['csv'])
with col2:
    file_tms_reg = st.file_uploader("2. TMS Regular Report (.csv)", type=['csv'])
with col3:
    file_tms_ffi = st.file_uploader("3. TMS FFI Report (.csv)", type=['csv'])

# Helper function to handle encoding issues gracefully
def load_csv(uploaded_file):
    try:
        # Try standard utf-8 encoding first
        return pd.read_csv(uploaded_file, encoding='utf-8')
    except UnicodeDecodeError:
        # Fallback to latin1 for files exported from certain legacy systems/Excel
        uploaded_file.seek(0)
        return pd.read_csv(uploaded_file, encoding='latin1')

# Helper function to apply conditional formatting to the DETAILED dataframe
def color_cells(row):
    styles = [''] * len(row)
    red_style = 'background-color: #ffcccc; color: #cc0000; font-weight: bold;'
    
    if row['Issue Type'] == 'Status Mismatch':
        styles[row.index.get_loc('Intella State')] = red_style
        styles[row.index.get_loc('TMS Status')] = red_style
    elif row['Issue Type'] == 'Offer Issued & Inactive > 60 Days':
        styles[row.index.get_loc('Last Updated Date')] = red_style
        styles[row.index.get_loc('Days Inactive')] = red_style
        
    return styles

# Helper function to apply conditional formatting to the SUMMARY dataframe
def style_summary_cells(row):
    styles = [''] * len(row)
    gray_style = 'background-color: #D3D3D3; font-weight: bold;'
    
    # 将责任人、Issue Type、Case Count三个格子全部标为浅灰色
    if row['Issue Type'] == 'Total Anomalies':
        styles[row.index.get_loc('Responsible Person')] = gray_style
        styles[row.index.get_loc('Issue Type')] = gray_style
        styles[row.index.get_loc('Case Count')] = gray_style
        
    return styles

if file_intella and file_tms_reg and file_tms_ffi:
    try:
        df_intella = load_csv(file_intella)
        df_tms_reg = load_csv(file_tms_reg)
        df_tms_ffi = load_csv(file_tms_ffi)
        
        tms_cols = ['Requisition Number', 'Requisition status']
        
        if not all(col in df_tms_reg.columns for col in tms_cols) or not all(col in df_tms_ffi.columns for col in tms_cols):
            st.error("Error: Both TMS Regular and TMS FFI reports must contain 'Requisition Number' and 'Requisition status' columns.")
        elif 'job_id' not in df_intella.columns:
            st.error("Error: Intella Report must contain a 'job_id' column.")
        else:
            with st.spinner('Processing data...'):
                # Step 1: Combine TMS files
                df_tms = pd.concat([df_tms_reg[tms_cols], df_tms_ffi[tms_cols]], ignore_index=True)
                df_tms = df_tms.drop_duplicates(subset=['Requisition Number'])
                
                # Step 2: Merge Intella with TMS
                mr = pd.merge(df_intella, df_tms, left_on='job_id', right_on='Requisition Number', how='left')
                
                # --- NEW STEP: Filter assignment_group == 'RAS Agent' ---
                if 'assignment_group' in mr.columns:
                    mr = mr[mr['assignment_group'].astype(str).str.strip() == 'RAS Agent']
                else:
                    st.warning("Warning: 'assignment_group' column not found in Intella report. Skipping RAS Agent filter.")
                
                required_cols = ['state', 'Requisition status', 'number', 'assigned_to', 'sys_updated_on']
                missing_cols = [col for col in required_cols if col not in mr.columns]
                
                if missing_cols:
                    st.error(f"Error: Missing columns in the merged data to perform anomaly checks: {', '.join(missing_cols)}")
                else:
                    mr['sys_updated_on_dt'] = pd.to_datetime(mr['sys_updated_on'], errors='coerce', dayfirst=True)
                    today = pd.Timestamp.today()
                    
                    anomalies = []
                    
                    # Step 3: Find anomaly cases
                    for index, row in mr.iterrows():
                        state = str(row['state']).strip() if pd.notna(row['state']) else ""
                        req_status = str(row['Requisition status']).strip() if pd.notna(row['Requisition status']) else ""
                        case_num = str(row['number'])
                        assigned_to = str(row['assigned_to'])
                        
                        def create_record(issue_type, intella_st="", tms_st="", last_updated="", days_inactive=""):
                            return {
                                'Responsible Person': assigned_to,
                                'Issue Type': issue_type,
                                'Case Number (JPR)': case_num,
                                'Intella State': intella_st,
                                'TMS Status': tms_st,
                                'Last Updated Date': last_updated,
                                'Days Inactive': days_inactive
                            }
                        
                        # Condition 1: Mismatched state and Requisition status
                        if state != req_status:
                            anomalies.append(create_record('Status Mismatch', state, req_status))
                            
                        # Condition 2: Either status is 'GSSC Process Completed'
                        if state == 'GSSC Process Completed' or req_status == 'GSSC Process Completed':
                            anomalies.append(create_record('Status is GSSC Process Completed', state, req_status))
                            
                        # Condition 3: Status is 'Offer issued' AND sys_updated_on is older than 60 days
                        if state == 'Offer issued' or req_status == 'Offer issued':
                            if pd.notna(row['sys_updated_on_dt']):
                                days_diff = (today - row['sys_updated_on_dt']).days
                                if days_diff > 60:
                                    formatted_date = row['sys_updated_on_dt'].strftime('%Y-%m-%d')
                                    anomalies.append(create_record('Offer Issued & Inactive > 60 Days', state, req_status, formatted_date, days_diff))
                                    
                    # Display results
                    if anomalies:
                        df_anomalies = pd.DataFrame(anomalies)
                        
                        # Sort detailed list
                        df_anomalies = df_anomalies.sort_values(
                            by=['Responsible Person', 'Issue Type'], 
                            ascending=[True, True]
                        ).reset_index(drop=True)
                        
                        # --- Create Standard Summary for UI ---
                        summary_df = df_anomalies.groupby(['Responsible Person', 'Issue Type']).agg(
                            Case_Count=('Case Number (JPR)', 'count'),
                            Specific_JPRs=('Case Number (JPR)', lambda x: ', '.join(x.unique()))
                        ).reset_index()
                        
                        # --- Create Formatted Summary for Excel (with Totals and Empty Rows) ---
                        formatted_summary_data = []
                        unique_persons_sorted = sorted(df_anomalies['Responsible Person'].unique())
                        
                        for person in unique_persons_sorted:
                            person_df = df_anomalies[df_anomalies['Responsible Person'] == person]
                            total_count = len(person_df)
                            
                            # 1. Add Total Row for the person
                            formatted_summary_data.append({
                                'Responsible Person': person,
                                'Issue Type': 'Total Anomalies',
                                'Case Count': total_count,
                                'Specific JPRs': ''
                            })
                            
                            # 2. Add Breakdown Rows
                            person_summary = summary_df[summary_df['Responsible Person'] == person]
                            for _, r in person_summary.iterrows():
                                formatted_summary_data.append({
                                    'Responsible Person': person,
                                    'Issue Type': f"  └ {r['Issue Type']}", # Add slight indent for readability
                                    'Case Count': r['Case_Count'],
                                    'Specific JPRs': r['Specific_JPRs']
                                })
                                
                            # 3. Add an Empty Row as separator
                            formatted_summary_data.append({
                                'Responsible Person': '',
                                'Issue Type': '',
                                'Case Count': '',
                                'Specific JPRs': ''
                            })
                            
                        # Convert to DataFrame (drop the very last empty row for neatness)
                        df_formatted_summary = pd.DataFrame(formatted_summary_data).iloc[:-1]
                        
                        # Apply gray styling to Summary DataFrame
                        styled_summary = df_formatted_summary.style.apply(style_summary_cells, axis=1)
                        
                        st.success(f"Processing complete! Found {len(df_anomalies)} anomaly records.")
                        
                        # --- Section: Anomaly Summary ---
                        st.subheader("📊 Anomaly Summary")
                        # Show styled summary in UI
                        st.dataframe(styled_summary, use_container_width=True)
                        
                        # Export Styled Summary to Excel
                        output_summary = io.BytesIO()
                        with pd.ExcelWriter(output_summary, engine='openpyxl') as writer:
                            # Use to_excel on the Styler object to preserve formatting
                            styled_summary.to_excel(writer, sheet_name='Anomaly Summary', index=False)
                            worksheet_sum = writer.sheets['Anomaly Summary']
                            
                            # Auto-fit columns for Summary
                            for idx, col in enumerate(df_formatted_summary.columns):
                                max_data_len = df_formatted_summary[col].astype(str).map(len).max() if not df_formatted_summary.empty else 0
                                col_len = max(max_data_len, len(str(col))) + 2
                                col_letter = get_column_letter(idx + 1)
                                worksheet_sum.column_dimensions[col_letter].width = col_len
                                
                        excel_summary_data = output_summary.getvalue()
                        
                        st.download_button(
                            label="📥 Download Summary (Excel)",
                            data=excel_summary_data,
                            file_name='anomaly_summary_report.xlsx',
                            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        )
                        
                        st.markdown("---")
                        
                        # --- Section: Detailed Anomaly List ---
                        st.subheader("🗂️ Detailed Anomaly List")
                        st.markdown("*Note: Highlighted colors are only visible in this online preview.*")
                        
                        styled_anomalies = df_anomalies.style.apply(color_cells, axis=1)
                        st.dataframe(styled_anomalies, use_container_width=True)
                        
                        # Export Detailed to Excel
                        output_detailed = io.BytesIO()
                        with pd.ExcelWriter(output_detailed, engine='openpyxl') as writer:
                            for person in unique_persons_sorted:
                                person_df = df_anomalies[df_anomalies['Responsible Person'] == person].copy()
                                styled_person_df = person_df.style.apply(color_cells, axis=1)
                                
                                safe_sheet_name = re.sub(r'[\\/*?:\[\]]', '', str(person))
                                safe_sheet_name = safe_sheet_name[:31] if safe_sheet_name else "Unknown"
                                
                                styled_person_df.to_excel(writer, sheet_name=safe_sheet_name, index=False)
                                
                                # Auto-fit columns for Detailed
                                worksheet_det = writer.sheets[safe_sheet_name]
                                for idx, col in enumerate(person_df.columns):
                                    max_data_len = person_df[col].astype(str).map(len).max() if not person_df.empty else 0
                                    col_len = max(max_data_len, len(str(col))) + 2
                                    col_letter = get_column_letter(idx + 1)
                                    worksheet_det.column_dimensions[col_letter].width = col_len
                        
                        excel_detailed_data = output_detailed.getvalue()
                        
                        st.download_button(
                            label="📥 Download Detailed List (Excel)",
                            data=excel_detailed_data,
                            file_name='anomaly_detailed_list.xlsx',
                            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        )
                        
                    else:
                        st.success("Processing complete! No anomaly cases were found based on the given rules.")
                        
    except Exception as e:
        st.error(f"An unexpected error occurred while processing the files: {e}")
