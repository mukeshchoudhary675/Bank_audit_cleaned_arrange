import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="State Splitter + Audit Tracking", layout="wide")
st.title("📊 Excel Cleaner with State-wise Sheets + Audit Tracking")

uploaded_file = st.file_uploader("📁 Upload your Excel file", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.success(f"✅ File uploaded. It has {df.shape[0]} rows and {df.shape[1]} columns.")

    st.subheader("🔗 Map Columns to Your Desired Structure")

    # Final desired columns (15)
    desired_columns = [
        "Warehouse_Code", "Warehouse_Name", "State", "Region", "Location", "CM_Name",
        "Customer_Name", "WHR/SR/ISIN_No", "Commodity_Name", "Commodity_Variety",
        "Balance_No_of_Bags", "OS_Quantit(MT)","Warehouse_Address", "Warehouse_Type", "CM_Location_Name", "Auditor"
    ]

    # Show dropdowns to map uploaded columns to desired format
    mapping = {}
    col1, col2 = st.columns(2)
    with col1:
        for col in desired_columns[:len(desired_columns) // 2]:
            mapping[col] = st.selectbox(f"Select column for: **{col}**", df.columns, key=col)
    with col2:
        for col in desired_columns[len(desired_columns) // 2:]:
            mapping[col] = st.selectbox(f"Select column for: **{col}**", df.columns, key=col)

    if st.button("🔄 Process and Generate Excel"):
        try:
            # Step 1: Reorder and rename columns
            selected_df = df[[mapping[col] for col in desired_columns]]
            selected_df.columns = desired_columns

            # Step 2: Build "Audit Tracking" sheet
            audit_cols = [
                "Warehouse_Code", "Warehouse_Name", "State", "Region", "CM_Name", "Location"
            ]
            audit_df = selected_df[audit_cols].drop_duplicates(subset=["Warehouse_Code"])
            audit_df.insert(0, "Sr No.", range(1, len(audit_df) + 1))

            # Step 3: Export everything to Excel
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                # First sheet: Audit Tracking
                audit_df.to_excel(writer, sheet_name="Audit Tracking", index=False)

                # State-wise sheets
                for state, group in selected_df.groupby("State"):
                    sheet_name = str(state)[:31] if pd.notna(state) else "Unknown"
                    group.to_excel(writer, sheet_name=sheet_name, index=False)

            st.success("✅ File processed successfully!")

            st.download_button(
                label="📥 Download Final Excel",
                data=output.getvalue(),
                file_name="statewise_with_audit_tracking.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"❌ Error while processing: {e}")
