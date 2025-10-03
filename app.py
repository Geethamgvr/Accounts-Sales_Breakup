import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

# Set page configuration
st.set_page_config(
    page_title="Sales Report Generator",
    page_icon="📊",
    layout="wide"
)

# App title and description
st.title("📊 Sales Report Generator")
st.markdown("Upload your CSV file to generate a formatted sales report with subtotals and grand totals.")

# File upload section
st.header("1. Upload Your Data")
uploaded_file = st.file_uploader("Choose a CSV file", type=['csv'])

# Add some information about expected file format
with st.expander("Expected file format"):
    st.markdown("""
    The application expects a CSV file with the following columns:
    - Online Reference Name
    - Table No
    - Order Type
    - Main Category
    - After Discount
    - CGST
    - SGST
    - Delivery Charge
    - Total Price
    
    The script will skip the first 5 rows of your CSV file.
    """)

if uploaded_file is not None:
    try:
        # Load the data
        df = pd.read_csv(uploaded_file, skiprows=5).iloc[:-1]
        
        # Show raw data preview
        st.header("2. Data Preview")
        st.dataframe(df.head())
        
        # Processing steps
        st.header("3. Processing Data")
        
        with st.spinner("Processing your data..."):
            # Filter Online Reference Name
            df['Online Reference Name'] = df['Online Reference Name'].astype(str).apply(
                lambda x: x if 'swiggy' in x.lower() or 'zomato' in x.lower() else '')

            # Classify Table No
            df['Table No'] = df['Table No'].astype(str).str.lower().apply(
                lambda x: 'Counter Sweet Sales' if x.startswith('sw') else 'Scrap Sales' if x.startswith('sr') else '')

            # Create Sub Category
            df['Sub Category'] = df.apply(
                lambda row: f"{row['Table No']} {row['Online Reference Name']}".strip() if row['Online Reference Name'] else row['Table No'],
                axis=1
            )

            # Group and aggregate
            grouped = df.groupby(['Order Type', 'Sub Category', 'Main Category']).agg({
                'After Discount': 'sum',
                'CGST': 'sum',
                'SGST': 'sum',
                'Delivery Charge': 'sum',
                'Total Price': 'sum'
            }).reset_index()

            # Define custom order for Order Type
            order_type_order = ['Dine-In', 'Take-Away', 'Delivery']
            grouped['Order Type'] = pd.Categorical(grouped['Order Type'], categories=order_type_order, ordered=True)
            grouped = grouped.sort_values(by=['Order Type', 'Sub Category', 'Main Category'])

            # Function to build final output with subtotals
            def build_final_table(df):
                result = []

                for order_type in order_type_order:
                    if order_type not in df['Order Type'].unique():
                        continue
                        
                    odf = df[df['Order Type'] == order_type]
                    order_type_written = False

                    for sub_cat in odf['Sub Category'].unique():
                        sdf = odf[odf['Sub Category'] == sub_cat]

                        for _, row in sdf.iterrows():
                            result.append({
                                'Order Type': order_type if not order_type_written else '',
                                'Sub Category': sub_cat,
                                'Main Category': row['Main Category'],
                                'After Discount': row['After Discount'],
                                'CGST': row['CGST'],
                                'SGST': row['SGST'],
                                'Delivery Charge': row['Delivery Charge'],
                                'Total Price': row['Total Price']
                            })
                            order_type_written = True

                        # Subcategory total
                        subtotal = sdf.select_dtypes(include='number').sum()
                        result.append({
                            'Order Type': '',
                            'Sub Category': f"{sub_cat} Total",
                            'Main Category': '',
                            'After Discount': subtotal['After Discount'],
                            'CGST': subtotal['CGST'],
                            'SGST': subtotal['SGST'],
                            'Delivery Charge': subtotal['Delivery Charge'],
                            'Total Price': subtotal['Total Price']
                        })

                    # Order Type total
                    order_total = odf.select_dtypes(include='number').sum()
                    result.append({
                        'Order Type': order_type,
                        'Sub Category': f"{order_type.strip()} Total",
                        'Main Category': '',
                        'After Discount': order_total['After Discount'],
                        'CGST': order_total['CGST'],
                        'SGST': order_total['SGST'],
                        'Delivery Charge': order_total['Delivery Charge'],
                        'Total Price': order_total['Total Price']
                    })
                    
                    # Add empty line after Take-Away section
                    if order_type == 'Take-Away':
                        result.append({
                            'Order Type': '',
                            'Sub Category': '',
                            'Main Category': '',
                            'After Discount': '',
                            'CGST': '',
                            'SGST': '',
                            'Delivery Charge': '',
                            'Total Price': ''
                        })

                final_df = pd.DataFrame(result)

                # Remove repeated values
                for col in ['Order Type', 'Sub Category']:
                    final_df[col] = final_df[col].where(final_df[col] != final_df[col].shift(), '')

                return final_df

            # Build the final report
            final = build_final_table(grouped)

            # Add Grand Total (only Delivery, Dine-In, Take-Away)
            totals_to_include = ['Delivery Total', 'Dine-In Total', 'Take-Away Total']
            grand_rows = final[final['Sub Category'].isin(totals_to_include)]

            grand_total = {
                'Order Type': 'Grand Total',
                'Sub Category': '',
                'Main Category': '',
                'After Discount': grand_rows['After Discount'].sum(),
                'CGST': grand_rows['CGST'].sum(),
                'SGST': grand_rows['SGST'].sum(),
                'Delivery Charge': grand_rows['Delivery Charge'].sum(),
                'Total Price': grand_rows['Total Price'].sum()
            }

            final = pd.concat([final, pd.DataFrame([grand_total])], ignore_index=True)
        
        st.success("✅ Data processed successfully!")
        
        # Show processed data
        st.header("4. Processed Report Preview")
        st.dataframe(final)
        
        # Download section
        st.header("5. Download Report")
        
        # Create a BytesIO buffer for the Excel file with formatting
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            final.to_excel(writer, index=False, sheet_name='Sales Report')
            
            # Get the xlsxwriter workbook and worksheet objects
            workbook = writer.book
            worksheet = writer.sheets['Sales Report']
            
            # Define formats
            header_format = workbook.add_format({
                'bold': True,
                'text_wrap': True,
                'valign': 'top',
                'fg_color': '#D7E4BC',
                'border': 1
            })
            
            total_format = workbook.add_format({
                'bold': True,
                'fg_color': '#FFE699',
                'border': 1
            })
            
            grand_total_format = workbook.add_format({
                'bold': True,
                'fg_color': '#F8CBAD',
                'border': 1
            })
            
            normal_format = workbook.add_format({'border': 1})
            
            # Apply header format
            for col_num, value in enumerate(final.columns.values):
                worksheet.write(0, col_num, value, header_format)
            
            # Apply formatting to data rows
            for row_num in range(1, len(final) + 1):
                sub_category = final.iloc[row_num-1]['Sub Category'] if row_num-1 < len(final) else ''
                order_type = final.iloc[row_num-1]['Order Type'] if row_num-1 < len(final) else ''
                
                # Determine the format based on content
                if 'Grand Total' in str(order_type):
                    cell_format = grand_total_format
                elif 'Total' in str(sub_category):
                    cell_format = total_format
                else:
                    cell_format = normal_format
                
                # Apply format to all cells in the row
                for col_num in range(len(final.columns)):
                    value = final.iloc[row_num-1, col_num] if row_num-1 < len(final) else ''
                    worksheet.write(row_num, col_num, value, cell_format)
            
            # Auto-adjust column widths
            for i, col in enumerate(final.columns):
                max_len = max(final[col].astype(str).str.len().max(), len(col)) + 2
                worksheet.set_column(i, i, max_len)
        
        # Create download button
        st.download_button(
            label="📥 Download Excel Report",
            data=output.getvalue(),
            file_name="Final_Sales_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"An error occurred: {str(e)}")
        st.info("Please make sure you've uploaded a properly formatted CSV file.")
else:
    st.info("👆 Please upload a CSV file to get started.")

# Add footer
st.markdown("---")
st.markdown("### 💡 Instructions")
st.markdown("""
1. Upload a CSV file with the expected format
2. The application will process your data
3. Preview the processed report
4. Download the final Excel file with highlighted totals
""")
