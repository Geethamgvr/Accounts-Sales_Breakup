import streamlit as st
import pandas as pd
import io

def process_files(csv_file, excel_file):
    """Process the uploaded CSV and Excel files"""
    
    # Read CSV
    df_csv = pd.read_csv(csv_file, skiprows=5)
    
    # Drop missing Order IDs, group, convert to int
    df_csv = df_csv.dropna(subset=["Order ID"])
    df_csv = df_csv.groupby("Order ID", as_index=False)["After Discount"].sum()
    df_csv["Order ID"] = df_csv["Order ID"].astype(int)
    
    # Read Excel
    df_excel = pd.read_excel(excel_file, skiprows=1)
    df_excel["After Discount"] = df_excel["OrderPrice"] - df_excel["Discount"]
    df_excel = df_excel[["OrderID", "After Discount", "Paymode"]]
    
    # Rename CSV column to match Excel's "OrderID" for comparison
    df_csv = df_csv.rename(columns={"Order ID": "OrderID"})
    
    # Find missing records
    missing_in_excel = df_csv[~df_csv["OrderID"].isin(df_excel["OrderID"])]
    missing_in_csv = df_excel[~df_excel["OrderID"].isin(df_csv["OrderID"])]
    
    return df_csv, df_excel, missing_in_excel, missing_in_csv

def main():
    st.set_page_config(page_title="Order Comparison Tool", layout="wide")
    
    st.title("📊 Order Comparison Tool")
    st.markdown("Upload CSV and Excel files to compare orders")
    
    # Create two columns for file uploads
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📁 Upload CSV File")
        csv_file = st.file_uploader("Choose CSV file", type=['csv'])
        if csv_file:
            st.success("✅ CSV file uploaded successfully!")
    
    with col2:
        st.subheader("📁 Upload Excel File")
        excel_file = st.file_uploader("Choose Excel file", type=['xlsx', 'xls'])
        if excel_file:
            st.success("✅ Excel file uploaded successfully!")
    
    # Process files when both are uploaded
    if csv_file and excel_file:
        try:
            with st.spinner("Processing files..."):
                df_csv, df_excel, missing_in_excel, missing_in_csv = process_files(csv_file, excel_file)
            
            # Display summary statistics
            st.subheader("📊 Summary Statistics")
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.metric("Total CSV Orders", len(df_csv))
            with col2:
                st.metric("Total Excel Orders", len(df_excel))
            with col3:
                st.metric("Missing in Excel", len(missing_in_excel), 
                         delta=f"${missing_in_excel['After Discount'].sum():.2f}")
            with col4:
                st.metric("Missing in CSV", len(missing_in_csv), 
                         delta=f"${missing_in_csv['After Discount'].sum():.2f}")
            
            # Display missing records
            tab1, tab2, tab3, tab4 = st.tabs(["📋 CSV Data", "📋 Excel Data", 
                                              "⚠️ Missing in Excel", "⚠️ Missing in CSV"])
            
            with tab1:
                st.subheader("CSV Data")
                st.dataframe(df_csv, use_container_width=True)
                
                # Download CSV data
                csv_buffer = io.StringIO()
                df_csv.to_csv(csv_buffer, index=False)
                st.download_button(
                    label="📥 Download CSV Data",
                    data=csv_buffer.getvalue(),
                    file_name="csv_data.csv",
                    mime="text/csv"
                )
            
            with tab2:
                st.subheader("Excel Data")
                st.dataframe(df_excel, use_container_width=True)
                
                # Download Excel data
                excel_buffer = io.BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                    df_excel.to_excel(writer, index=False, sheet_name='Excel Data')
                st.download_button(
                    label="📥 Download Excel Data",
                    data=excel_buffer.getvalue(),
                    file_name="excel_data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            
            with tab3:
                st.subheader(f"⚠️ OrderIDs in CSV but missing in Excel ({len(missing_in_excel)})")
                if len(missing_in_excel) > 0:
                    st.dataframe(missing_in_excel, use_container_width=True)
                    st.info(f"Total missing After Discount (CSV side): ${missing_in_excel['After Discount'].sum():.2f}")
                    
                    # Download missing data
                    csv_buffer = io.StringIO()
                    missing_in_excel.to_csv(csv_buffer, index=False)
                    st.download_button(
                        label="📥 Download Missing in Excel",
                        data=csv_buffer.getvalue(),
                        file_name="missing_in_excel.csv",
                        mime="text/csv"
                    )
                else:
                    st.success("🎉 No orders missing in Excel!")
            
            with tab4:
                st.subheader(f"⚠️ OrderIDs in Excel but missing in CSV ({len(missing_in_csv)})")
                if len(missing_in_csv) > 0:
                    st.dataframe(missing_in_csv, use_container_width=True)
                    st.info(f"Total missing After Discount (Excel side): ${missing_in_csv['After Discount'].sum():.2f}")
                    
                    # Download missing data
                    csv_buffer = io.StringIO()
                    missing_in_csv.to_csv(csv_buffer, index=False)
                    st.download_button(
                        label="📥 Download Missing in CSV",
                        data=csv_buffer.getvalue(),
                        file_name="missing_in_csv.csv",
                        mime="text/csv"
                    )
                else:
                    st.success("🎉 No orders missing in CSV!")
            
            # Full comparison report
            st.subheader("📝 Complete Comparison Report")
            
            # Create merged comparison
            merged = pd.merge(df_csv, df_excel, on='OrderID', how='outer', suffixes=('_CSV', '_Excel'))
            merged['Difference'] = merged['After Discount_CSV'] - merged['After Discount_Excel']
            
            st.dataframe(merged, use_container_width=True)
            
            # Download full report
            csv_buffer = io.StringIO()
            merged.to_csv(csv_buffer, index=False)
            st.download_button(
                label="📥 Download Complete Report",
                data=csv_buffer.getvalue(),
                file_name="complete_comparison_report.csv",
                mime="text/csv"
            )
            
        except Exception as e:
            st.error(f"❌ Error processing files: {str(e)}")
            st.info("Please make sure the files are in the correct format.")
    
    else:
        st.info("👈 Please upload both CSV and Excel files to begin analysis")
        
    # Instructions
    with st.expander("📖 Instructions"):
        st.markdown("""
        ### How to use this tool:
        1. **Upload CSV file** - Click on the first upload button and select your CSV file
        2. **Upload Excel file** - Click on the second upload button and select your Excel file
        3. **View Results** - The tool will automatically:
           - Process both files
           - Show summary statistics
           - Display missing records
           - Provide download options for all results
        
        ### File Format Requirements:
        - **CSV**: Should have columns 'Order ID' and 'After Discount' (after skipping 5 rows)
        - **Excel**: Should have columns 'OrderID', 'OrderPrice', 'Discount' (after skipping 1 row)
        
        ### What the tool does:
        - Compares orders between CSV and Excel files
        - Identifies orders missing in each file
        - Calculates totals for missing orders
        - Provides downloadable reports
        """)

if __name__ == "__main__":
    main()
