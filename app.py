import streamlit as st
import pandas as pd
import io

st.title("Order Data Comparison Tool")

# File uploaders
st.sidebar.header("Upload Files")
csv_file = st.sidebar.file_uploader("Upload BI.csv", type=["csv"])
excel_file = st.sidebar.file_uploader("Upload Local.xlsx", type=["xlsx"])

if csv_file and excel_file:
    try:
        # Read CSV — fixed skiprows to 5 (real header line)
        df_csv = pd.read_csv(csv_file, skiprows=5)

        # Drop missing Order IDs, group, convert to int
        df_csv = df_csv.dropna(subset=["Order Id"])
        df_csv = df_csv.groupby("Order ID", as_index=False)["After Discount"].sum()
        df_csv["Order ID"] = df_csv["Order ID"].astype(int)

        # Read Excel
        df_excel = pd.read_excel(excel_file, skiprows=1)
        df_excel["After Discount"] = df_excel["OrderPrice"] - df_excel["Discount"]
        df_excel = df_excel[["OrderID", "After Discount", "Paymode"]]

        # Rename CSV column to match Excel's "OrderID" for comparison
        df_csv = df_csv.rename(columns={"Order ID": "OrderID"})

        # ---- Find OrderIDs present in CSV but missing in Excel ----
        missing_in_excel = df_csv[~df_csv["OrderID"].isin(df_excel["OrderID"])]
        
        st.subheader("OrderIDs in CSV but missing in Excel:")
        st.dataframe(missing_in_excel)
        st.write("Total missing After Discount (CSV side):", missing_in_excel["After Discount"].sum())

        st.divider()

        # ---- Find OrderIDs present in Excel but missing in CSV ----
        missing_in_csv = df_excel[~df_excel["OrderID"].isin(df_csv["OrderID"])]
        st.subheader("OrderIDs in Excel but missing in CSV:")
        st.dataframe(missing_in_csv)
        st.write("Total missing After Discount (Excel side):", missing_in_csv["After Discount"].sum())

        # Export functionality
        st.sidebar.header("Export Results")
        
        # Export missing_in_excel
        if not missing_in_excel.empty:
            csv_missing_excel = missing_in_excel.to_csv(index=False)
            st.sidebar.download_button(
                label="Download CSV Missing in Excel",
                data=csv_missing_excel,
                file_name="missing_in_excel.csv",
                mime="text/csv"
            )
        
        # Export missing_in_csv
        if not missing_in_csv.empty:
            csv_missing_csv = missing_in_csv.to_csv(index=False)
            st.sidebar.download_button(
                label="Download Excel Missing in CSV",
                data=csv_missing_csv,
                file_name="missing_in_csv.csv",
                mime="text/csv"
            )
        
        # Export combined summary
        summary_data = {
            "Metric": [
                "OrderIDs in CSV but missing in Excel (Count)",
                "Total missing After Discount (CSV side)",
                "OrderIDs in Excel but missing in CSV (Count)",
                "Total missing After Discount (Excel side)"
            ],
            "Value": [
                len(missing_in_excel),
                missing_in_excel["After Discount"].sum(),
                len(missing_in_csv),
                missing_in_csv["After Discount"].sum()
            ]
        }
        summary_df = pd.DataFrame(summary_data)
        csv_summary = summary_df.to_csv(index=False)
        st.sidebar.download_button(
            label="Download Summary Report",
            data=csv_summary,
            file_name="comparison_summary.csv",
            mime="text/csv"
        )

    except Exception as e:
        st.error(f"Error processing files: {e}")
else:
    st.info("Please upload both CSV and Excel files from the sidebar to begin comparison.")
