import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Dynamic Excel Processor", layout="wide")
st.title("📊 Dynamic Excel Cleaner with Audit Tracking + Optional State Split")

uploaded_file = st.file_uploader("📁 Upload your Excel file", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    all_columns = df.columns.tolist()

    st.success(f"✅ Uploaded successfully. File has {df.shape[0]} rows and {df.shape[1]} columns.")

    st.subheader("📌 Select Columns for Final Output")
    selected_columns = st.multiselect(
        "✅ Choose columns to include in the main cleaned file",
        options=all_columns,
        default=all_columns
    )

    reorder_columns = st.multiselect(
        "🧩 Reorder selected columns (drag to arrange)",
        options=selected_columns,
        default=selected_columns
    )

    st.subheader("🛠️ Select Columns for Audit Tracking Sheet")
    audit_columns = st.multiselect(
        "📋 Choose columns to include in the Audit Tracking sheet",
        options=all_columns
    )

    state_column = st.selectbox(
        "🌍 Select 'State' column if available (or leave blank)",
        options=["None"] + all_columns,
        index=0
    )

    if st.button("🔄 Process and Generate Excel"):
        try:
            # --- Main cleaned dataframe
            final_df = df[reorder_columns]

            # --- Audit Tracking sheet
            audit_df = df[audit_columns].drop_duplicates()
            audit_df.insert(0, "Sr No.", range(1, len(audit_df) + 1))

            # --- Write to Excel
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                # Sheet 1: Audit Tracking
                audit_df.to_excel(writer, sheet_name="Audit Tracking", index=False)

                # State-wise or single output
                if state_column != "None":
                    for state, group in final_df.groupby(state_column):
                        sheet_name = str(state)[:31] if pd.notna(state) else "Unknown"
                        group.to_excel(writer, sheet_name=sheet_name, index=False)
                else:
                    final_df.to_excel(writer, sheet_name="Cleaned Data", index=False)

            st.success("✅ Excel file is ready!")

            st.download_button(
                label="📥 Download Excel File",
                data=output.getvalue(),
                file_name="dynamic_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"❌ Error while processing: {e}")
