pip install xlsxwriter

import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

# Page configuration
st.set_page_config(page_title="Sales Report Generator", page_icon="📊", layout="wide")
st.title("📊 Sales Report Generator")
st.markdown("Upload your CSV file to generate a formatted sales report with subtotals and grand totals.")

# File upload
uploaded_file = st.file_uploader("Choose a CSV file", type=['csv'])

if uploaded_file is not None:
    try:
        # Load and process data
        df = pd.read_csv(uploaded_file, skiprows=5).iloc[:-1]
        
        st.header("Data Preview")
        st.dataframe(df.head())
        
        with st.spinner("Processing your data..."):
            # Process data
            df['Online Reference Name'] = df['Online Reference Name'].astype(str).apply(
                lambda x: x if 'swiggy' in x.lower() or 'zomato' in x.lower() else '')
            
            df['Table No'] = df['Table No'].astype(str).str.lower().apply(
                lambda x: 'Counter Sweet Sales' if x.startswith('sw') else 'Scrap Sales' if x.startswith('sr') else '')
            
            df['Sub Category'] = df.apply(
                lambda row: f"{row['Table No']} {row['Online Reference Name']}".strip() if row['Online Reference Name'] else row['Table No'], axis=1)

            # Group and aggregate
            grouped = df.groupby(['Order Type', 'Sub Category', 'Main Category']).agg({
                'After Discount': 'sum', 'CGST': 'sum', 'SGST': 'sum', 
                'Delivery Charge': 'sum', 'Total Price': 'sum'
            }).reset_index().sort_values(by=['Order Type', 'Sub Category', 'Main Category'])

            # Build final table with subtotals
            result = []
            for order_type in grouped['Order Type'].unique():
                odf = grouped[grouped['Order Type'] == order_type]
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
                        'Order Type': '', 'Sub Category': f"{sub_cat} Total", 'Main Category': '',
                        'After Discount': subtotal['After Discount'], 'CGST': subtotal['CGST'],
                        'SGST': subtotal['SGST'], 'Delivery Charge': subtotal['Delivery Charge'],
                        'Total Price': subtotal['Total Price']
                    })

                # Order Type total
                order_total = odf.select_dtypes(include='number').sum()
                result.append({
                    'Order Type': order_type, 'Sub Category': f"{order_type.strip()} Total", 'Main Category': '',
                    'After Discount': order_total['After Discount'], 'CGST': order_total['CGST'],
                    'SGST': order_total['SGST'], 'Delivery Charge': order_total['Delivery Charge'],
                    'Total Price': order_total['Total Price']
                })

            final = pd.DataFrame(result)
            
            # Remove repeated values
            for col in ['Order Type', 'Sub Category']:
                final[col] = final[col].where(final[col] != final[col].shift(), '')
            
            # Add Grand Total - replace NaN with 0 to avoid issues
            totals_to_include = ['Delivery Total', 'Dine-In Total', 'Take-Away Total']
            grand_rows = final[final['Sub Category'].isin(totals_to_include)]
            
            # Fill NaN values with 0 before summing 
            numeric_cols = ['After Discount', 'CGST', 'SGST', 'Delivery Charge', 'Total Price']
            grand_rows[numeric_cols] = grand_rows[numeric_cols].fillna(0)
            
            grand_total = {
                'Order Type': 'Grand Total', 'Sub Category': '', 'Main Category': '',
                'After Discount': grand_rows['After Discount'].sum(),
                'CGST': grand_rows['CGST'].sum(), 'SGST': grand_rows['SGST'].sum(),
                'Delivery Charge': grand_rows['Delivery Charge'].sum(),
                'Total Price': grand_rows['Total Price'].sum()
            }
            
            # Add grand total directly without blank row
            final = pd.concat([final, pd.DataFrame([grand_total])], ignore_index=True)
            
            # Fill all NaN values 
            for col in final.columns:
                if final[col].dtype in ['float64', 'int64']:
                    final[col] = final[col].fillna(0)
                else:
                    final[col] = final[col].fillna('')

        st.success("✅ Data processed successfully!")
        
        # Display formatted preview
        st.header("Report Preview")
        
        def highlight_totals(row):
            # Subcategory totals (light yellow)
            if 'Total' in str(row['Sub Category']) and row['Sub Category'] not in ['Delivery Total', 'Dine-In Total', 'Take-Away Total']:
                return ['background-color: #fff2cc; font-weight: bold;'] * len(row)
            
            # Order Type totals - same color for all order types (light blue)
            elif row['Sub Category'] in ['Delivery Total', 'Dine-In Total', 'Take-Away Total']:
                return ['background-color: #b8cce4; font-weight: bold;'] * len(row)
            
            # Grand Total (dark orange)
            elif row['Order Type'] == 'Grand Total':
                return ['background-color: #c65911; color: white; font-weight: bold;'] * len(row)
            
            return [''] * len(row)
        
        styled_final = final.style.apply(highlight_totals, axis=1)
        st.dataframe(styled_final, height=400, use_container_width=True, hide_index=True)

        # Download Excel
        st.header("Download Report")
        output = BytesIO()
        
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            final.to_excel(writer, index=False, sheet_name='Sales Report', startrow=1)
            
            workbook = writer.book
            worksheet = writer.sheets['Sales Report']
            
            # Define formats
            header_format = workbook.add_format({
                'bold': True, 'text_wrap': True, 'valign': 'top', 'align': 'center',
                'fg_color': '#366092', 'font_color': 'white', 'border': 1
            })
            
            # Same format for all order type totals (light blue)
            ordertype_format = workbook.add_format({'bold': True, 'fg_color': '#b8cce4', 'border': 1, 'num_format': '#,##0.00'})
            subcat_format = workbook.add_format({'bold': True, 'fg_color': '#fff2cc', 'border': 1, 'num_format': '#,##0.00'})
            grand_format = workbook.add_format({'bold': True, 'fg_color': '#c65911', 'font_color': 'white', 'border': 1, 'num_format': '#,##0.00'})
            normal_format = workbook.add_format({'border': 1, 'num_format': '#,##0.00'})
            text_format = workbook.add_format({'border': 1})
            
            # Format headers
            for col_num, value in enumerate(final.columns.values):
                worksheet.write(1, col_num, value, header_format)
            
            # Format data rows
            for row_num in range(2, len(final) + 2):
                if row_num - 2 < len(final):
                    row_data = final.iloc[row_num - 2]
                    
                    for col_num in range(len(final.columns)):
                        cell_value = final.iloc[row_num - 2, col_num]
                        
                        # Determine the right format
                        if row_data['Order Type'] == 'Grand Total':
                            cell_format = grand_format
                        elif row_data['Sub Category'] in ['Delivery Total', 'Dine-In Total', 'Take-Away Total']:
                            cell_format = ordertype_format
                        elif 'Total' in str(row_data['Sub Category']):
                            cell_format = subcat_format
                        elif pd.api.types.is_number(cell_value):
                            cell_format = normal_format
                        else:
                            cell_format = text_format
                        
                        worksheet.write(row_num, col_num, cell_value, cell_format)
            
            # Auto-adjust column widths
            for i, col in enumerate(final.columns):
                max_len = max(final[col].astype(str).apply(len).max(), len(col)) + 2
                worksheet.set_column(i, i, max_len)
            
            # Add a title
            title_format = workbook.add_format({'bold': True, 'font_size': 16, 'align': 'center'})
            worksheet.merge_range(0, 0, 0, len(final.columns)-1, 'SALES BREAKUP REPORT', title_format)
            worksheet.freeze_panes(2, 0)

        st.download_button(
            "📥 Download Excel Report",
            output.getvalue(),
            "Final_Sales_Report.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"An error occurred: {str(e)}")

else:
    st.info("👆 Please upload a CSV file to get started.")
