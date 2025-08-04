import streamlit as st
import pandas as pd
import io

st.title("📊 Flexible Excel Cleaner + State-wise Splitter")

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.success("✅ File uploaded successfully!")

    st.subheader("🔗 Map Your Columns")

    # Desired final column structure
    desired_columns = [
        "Warehouse_Code", "Warehouse_Name", "State", "Region", "Location", "CM_Name",
        "Customer_Name", "WHR/SR/ISIN_No", "Commodity_Name", "Commodity_Variety",
        "Balance_No_of_Bags", "OS_Quantit(MT)","Warehouse_Address", "Warehouse_Type", "CM_Location_Name", "Auditor"
    ]

    # Create a mapping from desired columns to actual columns
    mapping = {}
    col1, col2 = st.columns(2)
    with col1:
        for col in desired_columns[:len(desired_columns)//2]:
            mapping[col] = st.selectbox(f"Select for **{col}**", df.columns, key=col)
    with col2:
        for col in desired_columns[len(desired_columns)//2:]:
            mapping[col] = st.selectbox(f"Select for **{col}**", df.columns, key=col)

    if st.button("🔄 Process and Split by State"):
        try:
            # Rename and reorder columns
            selected_df = df[[mapping[col] for col in desired_columns]]
            selected_df.columns = desired_columns  # Rename columns to standard names

            # Write to Excel with each State as a sheet
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                for state, group in selected_df.groupby("State"):
                sheet_name = str(state)[:31] if pd.notna(state) else "Unknown"
                group.to_excel(writer, sheet_name=sheet_name, index=False)


            st.success("✅ File processed successfully!")

            st.download_button(
                label="📥 Download State-wise Excel",
                data=output.getvalue(),
                file_name="state_wise_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"❌ Error while processing: {e}")
