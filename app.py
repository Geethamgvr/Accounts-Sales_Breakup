import streamlit as st
import pandas as pd
import io
import re

def find_column(df, possible_names):
    """Find a column by trying multiple possible names"""
    for name in possible_names:
        if name in df.columns:
            return name
    # Try case-insensitive matching
    for col in df.columns:
        if col.lower() in [n.lower() for n in possible_names]:
            return col
    return None

def process_files(csv_file, excel_file):
    """Process the uploaded CSV and Excel files"""
    
    # Read CSV with error handling
    try:
        df_csv = pd.read_csv(csv_file, skiprows=5)
    except Exception as e:
        raise Exception(f"Error reading CSV file: {str(e)}")
    
    # Find the correct column names
    order_id_col = find_column(df_csv, ['Order ID', 'OrderID', 'Order_Id', 'Order', 'OrderId'])
    discount_col = find_column(df_csv, ['After Discount', 'AfterDiscount', 'Discount', 'Amount'])
    
    if order_id_col is None:
        st.error(f"Available CSV columns: {list(df_csv.columns)}")
        raise Exception("Could not find 'Order ID' column in CSV file")
    
    if discount_col is None:
        st.error(f"Available CSV columns: {list(df_csv.columns)}")
        raise Exception("Could not find 'After Discount' column in CSV file")
    
    # Drop missing Order IDs
    df_csv = df_csv.dropna(subset=[order_id_col])
    
    # Group by Order ID and sum After Discount
    df_csv = df_csv.groupby(order_id_col, as_index=False)[discount_col].sum()
    
    # Convert Order ID to int if possible
    try:
        df_csv[order_id_col] = df_csv[order_id_col].astype(int)
    except:
        pass  # Keep as is if conversion fails
    
    # Read Excel
    try:
        df_excel = pd.read_excel(excel_file, skiprows=1)
    except Exception as e:
        raise Exception(f"Error reading Excel file: {str(e)}")
    
    # Find Excel columns
    excel_order_col = find_column(df_excel, ['OrderID', 'Order ID', 'Order_Id', 'Order', 'OrderId'])
    excel_price_col = find_column(df_excel, ['OrderPrice', 'Price', 'Amount', 'Order Amount'])
    excel_discount_col = find_column(df_excel, ['Discount', 'Discount Amount'])
    excel_paymode_col = find_column(df_excel, ['Paymode', 'Payment Mode', 'PaymentMethod', 'Payment'])
    
    if excel_order_col is None:
        st.error(f"Available Excel columns: {list(df_excel.columns)}")
        raise Exception("Could not find 'OrderID' column in Excel file")
    
    if excel_price_col is None:
        st.error(f"Available Excel columns: {list(df_excel.columns)}")
        raise Exception("Could not find 'OrderPrice' column in Excel file")
    
    # Calculate After Discount
    if excel_discount_col is not None:
        df_excel["After Discount"] = df_excel[excel_price_col] - df_excel[excel_discount_col]
    else:
        # If no discount column, just use the price
        df_excel["After Discount"] = df_excel[excel_price_col]
    
    # Select columns for Excel
    excel_columns = [excel_order_col, "After Discount"]
    if excel_paymode_col is not None:
        excel_columns.append(excel_paymode_col)
    
    df_excel = df_excel[excel_columns]
    
    # Rename columns for comparison
    df_csv = df_csv.rename(columns={order_id_col: "OrderID", discount_col: "After Discount"})
    df_excel = df_excel.rename(columns={excel_order_col: "OrderID"})
    
    # Handle different data types for comparison
    try:
        df_csv["OrderID"] = df_csv["OrderID"].astype(str).str.strip()
        df_excel["OrderID"] = df_excel["OrderID"].astype(str).str.strip()
    except:
        pass
    
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
                if len(missing_in_excel) > 0:
                    st.metric("Missing in Excel", len(missing_in_excel), 
                             delta=f"${missing_in_excel['After Discount'].sum():,.2f}")
                else:
                    st.metric("Missing in Excel", 0)
            with col4:
                if len(missing_in_csv) > 0:
                    st.metric("Missing in CSV", len(missing_in_csv), 
                             delta=f"${missing_in_csv['After Discount'].sum():,.2f}")
                else:
                    st.metric("Missing in CSV", 0)
            
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
                    st.info(f"Total missing After Discount (CSV side): ${missing_in_excel['After Discount'].sum():,.2f}")
                    
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
                    st.info(f"Total missing After Discount (Excel side): ${missing_in_csv['After Discount'].sum():,.2f}")
                    
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
            
            try:
                # Create merged comparison
                merged = pd.merge(df_csv, df_excel, on='OrderID', how='outer', suffixes=('_CSV', '_Excel'))
                
                # Handle missing values in difference calculation
                merged['After Discount_CSV'] = merged['After Discount_CSV'].fillna(0)
                merged['After Discount_Excel'] = merged['After Discount_Excel'].fillna(0)
                merged['Difference'] = merged['After Discount_CSV'] - merged['After Discount_Excel']
                
                # Add status column
                merged['Status'] = merged.apply(
                    lambda row: '✅ Match' if row['After Discount_CSV'] == row['After Discount_Excel'] 
                    else '⚠️ Difference' if row['After Discount_CSV'] != 0 and row['After Discount_Excel'] != 0
                    else '❌ Missing in CSV' if row['After Discount_CSV'] == 0
                    else '❌ Missing in Excel',
                    axis=1
                )
                
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
                st.warning(f"Could not create comparison report: {str(e)}")
            
        except Exception as e:
            st.error(f"❌ Error processing files: {str(e)}")
            st.info("💡 Please make sure the files are in the correct format.")
            
            # Show expected format
            with st.expander("📖 Expected File Format"):
                st.markdown("""
                ### CSV File Format:
                - Should have columns similar to: `Order ID`, `After Discount`
                - The first 5 rows are skipped
                
                ### Excel File Format:
                - Should have columns similar to: `OrderID`, `OrderPrice`, `Discount`
                - The first 1 row is skipped
                
                ### The tool will try to find columns with similar names automatically
                """)
    
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
        
        ### Features:
        - **Auto-detection** of column names (case-insensitive)
        - **Flexible column naming** - works with different variations
        - **Detailed reports** with download options
        - **Visual metrics** for quick insights
        
        ### What the tool does:
        - Compares orders between CSV and Excel files
        - Identifies orders missing in each file
        - Calculates totals for missing orders
        - Provides downloadable reports in CSV format
        """)

if __name__ == "__main__":
    main()
