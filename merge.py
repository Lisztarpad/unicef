import streamlit as st
import pandas as pd
import io

# Set page configuration
st.set_page_config(page_title="Smart Data Merge Tool", page_icon="🗂️", layout="wide")

st.title("🗂️ Smart Data Merge Tool")
st.markdown("Upload two tabular files, configure the merge rules, and export the merged result.")


# Cache data loading to prevent reloading on every interaction
@st.cache_data
def load_data(uploaded_file):
    try:
        if uploaded_file.name.endswith('.csv'):
            try:
                return pd.read_csv(uploaded_file, encoding='utf-8')
            except UnicodeDecodeError:
                uploaded_file.seek(0)
                return pd.read_csv(uploaded_file, encoding='gbk')
        elif uploaded_file.name.endswith(('.xls', '.xlsx')):
            return pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"Error reading file {uploaded_file.name}: {e}")
        return None
    return None


# --- Section 1: File Upload & Preview ---
col1, col2 = st.columns(2)

with col1:
    st.header("📄 Table A (Left)")
    file_a = st.file_uploader("Upload first file (CSV or Excel)", type=["csv", "xlsx", "xls"], key="file_a")
    if file_a:
        df_a = load_data(file_a)
        if df_a is not None:
            st.success(f"Successfully loaded {file_a.name} (Rows: {df_a.shape[0]}, Columns: {df_a.shape[1]})")
            st.dataframe(df_a.head(), use_container_width=True)

with col2:
    st.header("📄 Table B (Right)")
    file_b = st.file_uploader("Upload second file (CSV or Excel)", type=["csv", "xlsx", "xls"], key="file_b")
    if file_b:
        df_b = load_data(file_b)
        if df_b is not None:
            st.success(f"Successfully loaded {file_b.name} (Rows: {df_b.shape[0]}, Columns: {df_b.shape[1]})")
            st.dataframe(df_b.head(), use_container_width=True)

# --- Section 2: Merge Configurations ---
if file_a and file_b and (df_a is not None) and (df_b is not None):
    st.divider()
    st.header("⚙️ Merge Rule Settings")

    # Extract column names
    cols_a = df_a.columns.tolist()
    cols_b = df_b.columns.tolist()

    # Merge methods
    how_options = {
        "Inner Join - Keep rows that match in both tables": "inner",
        "Left Join - Keep all rows from Left Table": "left",
        "Right Join - Keep all rows from Right Table": "right",
        "Outer Join - Keep all rows from both tables": "outer"
    }

    col_settings_1, col_settings_2 = st.columns([1, 1])

    with col_settings_1:
        selected_how = st.selectbox("1. Select Merge Method (How)", list(how_options.keys()))
        merge_how = how_options[selected_how]

    with col_settings_2:
        match_type = st.radio("2. Match Key Type",
                              ["Same Column Names (On)", "Different Column Names (Left On / Right On)"],
                              horizontal=True)

    # Setup Merge Keys
    merge_kwargs = {}
    can_merge = True

    if match_type == "Same Column Names (On)":
        common_cols = list(set(cols_a) & set(cols_b))
        if not common_cols:
            st.warning(
                "⚠️ The two tables have no common columns. Please switch to 'Different Column Names' to manually specify the keys.")
            can_merge = False
        else:
            selected_on = st.multiselect("Select Merge Key(s)", common_cols,
                                         default=common_cols[0] if common_cols else None)
            if not selected_on:
                st.warning("Please select at least one merge key.")
                can_merge = False
            else:
                merge_kwargs["on"] = selected_on
    else:
        k_col1, k_col2 = st.columns(2)
        with k_col1:
            left_on = st.multiselect("Table A Key (Left On)", cols_a, max_selections=1)
        with k_col2:
            right_on = st.multiselect("Table B Key (Right On)", cols_b, max_selections=1)

        if not left_on or not right_on:
            st.warning("Please select merge keys for both Table A and Table B.")
            can_merge = False
        else:
            merge_kwargs["left_on"] = left_on[0]
            merge_kwargs["right_on"] = right_on[0]

    # Advanced Settings: Suffixes
    with st.expander("🛠️ Advanced Settings: Suffixes for overlapping columns"):
        c1, c2 = st.columns(2)
        with c1:
            suffix_x = st.text_input("Table A Suffix", value="_left")
        with c2:
            suffix_y = st.text_input("Table B Suffix", value="_right")
        merge_kwargs["suffixes"] = (suffix_x, suffix_y)

    # --- Section 3: Execution & Export ---
    st.divider()
    if can_merge:
        if st.button("🚀 Execute Merge", type="primary", use_container_width=True):
            try:
                # Execute Pandas Merge
                df_merged = pd.merge(df_a, df_b, how=merge_how, **merge_kwargs)

                st.success(
                    f"✅ Merge successful! The merged table has {df_merged.shape[0]} rows and {df_merged.shape[1]} columns.")

                # Result Preview
                st.subheader("👀 Preview of Merged Result (Top 50 rows)")
                st.dataframe(df_merged.head(50), use_container_width=True)

                # Export Settings
                st.subheader("💾 Export Merged Result")
                export_col1, export_col2 = st.columns(2)

                # Export as CSV (using utf-8-sig for Excel compatibility)
                csv = df_merged.to_csv(index=False).encode('utf-8-sig')
                with export_col1:
                    st.download_button(
                        label="📥 Download CSV File",
                        data=csv,
                        file_name="merged_data.csv",
                        mime="text/csv",
                        use_container_width=True
                    )

                # Export as Excel
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_merged.to_excel(writer, index=False, sheet_name='Merged_Data')
                excel_data = output.getvalue()

                with export_col2:
                    st.download_button(
                        label="📥 Download Excel File",
                        data=excel_data,
                        file_name="merged_data.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )

            except Exception as e:
                st.error(f"❌ Merge failed. Error details: {str(e)}")
                st.info(
                    "💡 Tip: If it's a type error, ensure the selected key columns have the same data type in both tables (e.g., both are text or both are numbers).")