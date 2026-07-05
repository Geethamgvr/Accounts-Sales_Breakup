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
        
        # Debug: Show CSV columns
        st.write("CSV Columns:", df_csv.columns.tolist())
        
        # Find the correct column names (case insensitive)
        csv_order_col = None
        csv_discount_col = None
        
        for col in df_csv.columns:
            if 'order' in col.lower() and 'id' in col.lower():
                csv_order_col = col
            if 'discount' in col.lower() and 'after' in col.lower():
                csv_discount_col = col
        
        # If not found, try alternative names
        if csv_order_col is None:
            for col in df_csv.columns:
                if 'order' in col.lower():
                    csv_order_col = col
                    break
        
        if csv_discount_col is None:
            for col in df_csv.columns:
                if 'discount' in col.lower():
                    csv_discount_col = col
                    break
        
        if csv_order_col is None or csv_discount_col is None:
            st.error(f"Could not find required columns. Available columns: {df_csv.columns.tolist()}")
            st.stop()
        
        # Rename columns for consistency
        df_csv = df_csv.rename(columns={csv_order_col: "Order ID", csv_discount_col: "After Discount"})

        # Drop missing Order IDs, group, convert to int
        df_csv = df_csv.dropna(subset=["Order ID"])
        df_csv = df_csv.groupby("Order ID", as_index=False)["After Discount"].sum()
        df_csv["Order ID"] = df_csv["Order ID"].astype(int)

        # Read Excel
        df_excel = pd.read_excel(excel_file, skiprows=1)
        
        # Debug: Show Excel columns
        st.write("Excel Columns:", df_excel.columns.tolist())
        
        # Find Excel columns
        excel_order_col = None
        excel_price_col = None
        excel_discount_col = None
        excel_paymode_col = None
        
        for col in df_excel.columns:
            if 'order' in col.lower() and 'id' in col.lower():
                excel_order_col = col
            if 'price' in col.lower() or 'orderprice' in col.lower():
                excel_price_col = col
            if 'discount' in col.lower():
                excel_discount_col = col
            if 'paymode' in col.lower() or 'pay mode' in col.lower():
                excel_paymode_col = col
        
        if excel_order_col is None or excel_price_col is None or excel_discount_col is None:
            st.error(f"Could not find required columns in Excel. Available columns: {df_excel.columns.tolist()}")
            st.stop()
        
        # Use found column names
        df_excel["After Discount"] = df_excel[excel_price_col] - df_excel[excel_discount_col]
        
        # Select columns
        cols_to_keep = [excel_order_col, "After Discount"]
        if excel_paymode_col:
            cols_to_keep.append(excel_paymode_col)
        
        df_excel = df_excel[cols_to_keep]
        df_excel = df_excel.rename(columns={excel_order_col: "OrderID"})

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
        st.write("Debug - CSV columns:", df_csv.columns.tolist() if 'df_csv' in locals() else "CSV not loaded")
        st.write("Debug - Excel columns:", df_excel.columns.tolist() if 'df_excel' in locals() else "Excel not loaded")
else:
    st.info("Please upload both CSV and Excel files from the sidebar to begin comparison.")
