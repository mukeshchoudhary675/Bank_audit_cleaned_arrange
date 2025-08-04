import streamlit as st
import pandas as pd
import io

# Desired column order
desired_columns = [
    "Warehouse_Code", "Warehouse_Name", "State", "Region", "Location", "CM_Name",
    "Customer_Name", "WHR/SR/ISIN_No", "Commodity_Name", "Commodity_Variety",
    "Balance_No_of_Bags", "OS_Quantit(MT)", "Warehouse_Type", "CM_Location_Name", "Auditor"
]

st.title("🧾 State-wise Excel Cleaner & Splitter")

uploaded_file = st.file_uploader("Upload Raw Excel File", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)

        # Check for missing columns
        missing_cols = [col for col in desired_columns if col not in df.columns]
        if missing_cols:
            st.error(f"These columns are missing in uploaded file: {missing_cols}")
        else:
            # Reorder columns
            df = df[desired_columns]

            # Group by 'State' and write each group to a sheet
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                for state, group in df.groupby('State'):
                    safe_state = str(state)[:31]  # Sheet name max length
                    group.to_excel(writer, sheet_name=safe_state, index=False)
                writer.save()

            st.success("✅ File processed successfully!")

            st.download_button(
                label="📥 Download State-wise Excel",
                data=output.getvalue(),
                file_name="state_wise_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"❌ Error: {e}")
