import pandas as pd
from datetime import time
import streamlit as st
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows

st.set_page_config(page_title="Bill-wise Item Sales", layout="wide")

# Custom CSS for colors and auto-fit
st.markdown("""
<style>
    /* Captain name rows - light blue background */
    .captain-row {
        background-color: #e6f3ff !important;
        font-weight: 500 !important;
    }
    /* Grand total row - light green background */
    .grandtotal-row {
        background-color: #d4edda !important;
        font-weight: bold !important;
    }
    /* Item rows - distinct colors per captain */
    .captain-0-items {
        background-color: #e6ffe6 !important;  /* Light green for first captain */
    }
    .captain-1-items {
        background-color: #e6f0ff !important;  /* Light blue for second captain */
    }
    .captain-2-items {
        background-color: #fff2e6 !important;  /* Light orange for third captain */
    }
    .captain-3-items {
        background-color: #ffe6f0 !important;  /* Light pink for fourth captain */
    }
    .captain-4-items {
        background-color: #f0e6ff !important;  /* Light purple for fifth captain */
    }
    /* Auto-fit columns */
    .dataframe-container {
        font-size: 14px;
        width: 100%;
        overflow-x: auto;
    }
    /* Summary section styling */
    .item-summary {
        background-color: #f8f9fa;
        padding: 20px;
        border-radius: 10px;
        margin: 10px 0;
    }
    .item-card {
        background-color: white;
        padding: 10px;
        border-radius: 5px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        margin: 5px;
        text-align: center;
    }
</style>
""", unsafe_allow_html=True)

st.title("📊 Bill-wise Item Sales Processor")
st.write("Upload your CSV file to generate the report")

# File upload
uploaded_file = st.file_uploader("Choose a CSV file", type="csv")

if uploaded_file is not None:
    try:
        # Read the uploaded file
        df = pd.read_csv(uploaded_file, skiprows=5)
        st.success("File uploaded successfully!")
        
        with st.spinner("Processing..."):
            # Filter data
            file1 = df.iloc[:-1] if len(df) > 0 else df
            file2 = file1[file1['Order Type'] == 'Dine-In'].copy()
            file = file2[['Item Name', 'Order Type', 'Quantity', 'Captain Name','Billed Time']]
            file = file[file['Captain Name'].str.lower() != 'without_captain'].copy()
            
            # Convert Billed Time
            file['Billed Time'] = pd.to_datetime(file['Billed Time'])
            
            # Time categorization
            def get_time_category(billed_time):
                t = billed_time.time()
                if time(6, 0) <= t <= time(11, 00):
                    return 'Breakfast'
                elif time(11, 1) <= t <= time(16, 0):
                    return 'Lunch'
                elif time(16, 1) <= t <= time(19, 0):
                    return 'Snacks'
                else:
                    return 'Late Night'
            
            file['Time Category'] = file['Billed Time'].apply(get_time_category)
            
            # Group items
            def Grouped_items(ItemName):
                item_name = str(ItemName).lower()
                if "tamilnadu meals" in item_name:
                    return "Tamilnadu Meals"
                elif 'curd rice' in item_name:
                    return "Classic Curd Rice"
                elif 'thali' in item_name:
                    return "North Indian Thali"
                elif "special soup" in item_name:
                    return "Day Spl Soup"
                elif "tandoori" in item_name:
                    return "Tandoori Platter"
                elif "kunafa" in item_name:
                    return 'Kunafa_item'
                elif "fried rice" in item_name:
                    return 'Fried rice'
                elif "noodles" in item_name:
                    return 'Noodles_Item'
                elif "falooda" in item_name:
                    return "Falooda_Item"
                else:
                    return None
            
            file['Item Name'] = file['Item Name'].apply(Grouped_items)
            file = file.dropna(subset=['Item Name']).copy()
            
            # Create report
            unique_captains = sorted(file['Captain Name'].unique())
            all_results = []
            
            # Track grand totals
            grand_total_breakfast = 0
            grand_total_lunch = 0
            grand_total_snacks = 0
            grand_total_latenight = 0
            grand_total_all = 0
            
            # Track item-wise totals
            item_totals = {}
            
            # Track which items belong to which captain for coloring
            captain_item_mapping = {}
            
            for captain in unique_captains:
                captain_data = file[file['Captain Name'] == captain]
                captain_items = sorted(captain_data['Item Name'].unique())
                captain_item_mapping[captain] = captain_items
                
                # First item for this captain - show captain name
                first_item = True
                
                for item in captain_items:
                    item_data = captain_data[captain_data['Item Name'] == item]
                    pivot = pd.pivot_table(
                        item_data,
                        values='Quantity',
                        columns='Time Category',
                        aggfunc='sum',
                        fill_value=0
                    )
                    
                    # Show captain name only for first item of each captain
                    if first_item:
                        captain_display = captain
                        first_item = False
                    else:
                        captain_display = ''
                    
                    # Get values - replace 0 with empty string for display
                    breakfast = int(pivot.get('Breakfast', 0))
                    lunch = int(pivot.get('Lunch', 0))
                    snacks = int(pivot.get('Snacks', 0))
                    late_night = int(pivot.get('Late Night', 0))
                    total = breakfast + lunch + snacks + late_night
                    
                    # Add to grand totals
                    grand_total_breakfast += breakfast
                    grand_total_lunch += lunch
                    grand_total_snacks += snacks
                    grand_total_latenight += late_night
                    grand_total_all += total
                    
                    # Add to item-wise totals
                    if item not in item_totals:
                        item_totals[item] = 0
                    item_totals[item] += total
                    
                    # Store as integers but will display blanks for zeros
                    row = {
                        'Captain Name': captain_display,
                        'Item Name': item,
                        'Breakfast': breakfast if breakfast > 0 else '',
                        'Lunch': lunch if lunch > 0 else '',
                        'Snacks': snacks if snacks > 0 else '',
                        'Late Night': late_night if late_night > 0 else '',
                        'Total': total
                    }
                    all_results.append(row)
            
            # Add grand total row at the bottom
            grand_total_row = {
                'Captain Name': 'GRAND TOTAL',
                'Item Name': '--- ALL ITEMS TOTAL ---',
                'Breakfast': grand_total_breakfast if grand_total_breakfast > 0 else '',
                'Lunch': grand_total_lunch if grand_total_lunch > 0 else '',
                'Snacks': grand_total_snacks if grand_total_snacks > 0 else '',
                'Late Night': grand_total_latenight if grand_total_latenight > 0 else '',
                'Total': grand_total_all
            }
            all_results.append(grand_total_row)
            
            # Final dataframe
            final_result = pd.DataFrame(all_results)
            
            # Display results with colors and auto-fit
            st.subheader("📋 Processed Results")
            
            # Create color mapping function for display
            def color_rows(row):
                if row['Captain Name'] == 'GRAND TOTAL':
                    return ['background-color: #d4edda; font-weight: bold'] * len(row)
                elif row['Captain Name'] != '' and row['Captain Name'] != 'GRAND TOTAL':
                    # This is a captain header row
                    return ['background-color: #e6f3ff; font-weight: 500'] * len(row)
                else:
                    # This is an item row - find which captain it belongs to
                    for captain, items in captain_item_mapping.items():
                        if row['Item Name'] in items:
                            captain_index = list(captain_item_mapping.keys()).index(captain)
                            color_class = f'captain-{(captain_index % 5)}-items'
                            if color_class == 'captain-0-items':
                                return ['background-color: #e6ffe6'] * len(row)  # Light green
                            elif color_class == 'captain-1-items':
                                return ['background-color: #e6f0ff'] * len(row)  # Light blue
                            elif color_class == 'captain-2-items':
                                return ['background-color: #fff2e6'] * len(row)  # Light orange
                            elif color_class == 'captain-3-items':
                                return ['background-color: #ffe6f0'] * len(row)  # Light pink
                            elif color_class == 'captain-4-items':
                                return ['background-color: #f0e6ff'] * len(row)  # Light purple
                            break
                return [''] * len(row)
            
            # Apply styling
            styled_df = final_result.style.apply(color_rows, axis=1)
            
            # Display with auto-fit columns
            st.dataframe(styled_df, use_container_width=True, height=500)
            
            # Item-wise summary
            st.subheader("📦 Item-wise Total Quantity")
            
            # Create item summary in a grid
            item_summary_df = pd.DataFrame([
                {'Item': item, 'Total Quantity': qty} 
                for item, qty in sorted(item_totals.items(), key=lambda x: x[1], reverse=True)
            ])
            
            # Display items in multiple columns
            cols = st.columns(3)
            for idx, row in item_summary_df.iterrows():
                col_idx = idx % 3
                with cols[col_idx]:
                    st.markdown(f"""
                    <div class="item-card">
                        <strong>{row['Item']}</strong><br>
                        <span style="font-size: 24px; color: #FF4B4B;">{row['Total Quantity']}</span>
                    </div>
                    """, unsafe_allow_html=True)
            
            # Time-wise summary
            st.subheader("⏰ Time-wise Summary")
            col1, col2, col3, col4, col5 = st.columns(5)
            
            with col1:
                st.metric("Total Items", grand_total_all, delta=None)
            with col2:
                st.metric("Breakfast", grand_total_breakfast, delta=None)
            with col3:
                st.metric("Lunch", grand_total_lunch, delta=None)
            with col4:
                st.metric("Snacks", grand_total_snacks, delta=None)
            with col5:
                st.metric("Late Night", grand_total_latenight, delta=None)
            
            # Download buttons with Excel formatting
            st.subheader("💾 Download")
            col1, col2 = st.columns(2)
            
            with col1:
                # Create formatted Excel file
                output = BytesIO()
                
                # Create a workbook and add a worksheet
                wb = Workbook()
                ws = wb.active
                ws.title = "Sales Data"
                
                # Define colors (RGB hex to Excel fill)
                colors = {
                    'captain_header': PatternFill(start_color='E6F3FF', end_color='E6F3FF', fill_type='solid'),
                    'grand_total': PatternFill(start_color='D4EDDA', end_color='D4EDDA', fill_type='solid'),
                    'captain_0': PatternFill(start_color='E6FFE6', end_color='E6FFE6', fill_type='solid'),  # Light green
                    'captain_1': PatternFill(start_color='E6F0FF', end_color='E6F0FF', fill_type='solid'),  # Light blue
                    'captain_2': PatternFill(start_color='FFF2E6', end_color='FFF2E6', fill_type='solid'),  # Light orange
                    'captain_3': PatternFill(start_color='FFE6F0', end_color='FFE6F0', fill_type='solid'),  # Light pink
                    'captain_4': PatternFill(start_color='F0E6FF', end_color='F0E6FF', fill_type='solid'),  # Light purple
                }
                
                # Add headers
                headers = ['Captain Name', 'Item Name', 'Breakfast', 'Lunch', 'Snacks', 'Late Night', 'Total']
                for col_num, header in enumerate(headers, 1):
                    cell = ws.cell(row=1, column=col_num)
                    cell.value = header
                    cell.font = Font(bold=True)
                    cell.alignment = Alignment(horizontal='center')
                    cell.border = Border(
                        left=Side(style='thin'), 
                        right=Side(style='thin'),
                        top=Side(style='thin'), 
                        bottom=Side(style='thin')
                    )
                
                # Add data with colors
                for row_num, row_data in enumerate(all_results, 2):
                    # Determine row color
                    if row_data['Captain Name'] == 'GRAND TOTAL':
                        row_color = colors['grand_total']
                        font = Font(bold=True)
                    elif row_data['Captain Name'] != '' and row_data['Captain Name'] != 'GRAND TOTAL':
                        row_color = colors['captain_header']
                        font = Font(bold=False)
                    else:
                        # Find which captain this item belongs to
                        row_color = colors['captain_0']  # default to light green
                        for captain, items in captain_item_mapping.items():
                            if row_data['Item Name'] in items:
                                captain_index = list(captain_item_mapping.keys()).index(captain)
                                color_key = f'captain_{(captain_index % 5)}'
                                row_color = colors.get(color_key, colors['captain_0'])
                                break
                        font = Font(bold=False)
                    
                    for col_num, col_name in enumerate(headers, 1):
                        cell = ws.cell(row=row_num, column=col_num)
                        cell.value = row_data[col_name]
                        cell.fill = row_color
                        cell.font = font
                        cell.alignment = Alignment(horizontal='center' if col_name in ['Breakfast', 'Lunch', 'Snacks', 'Late Night', 'Total'] else 'left')
                        cell.border = Border(
                            left=Side(style='thin'), 
                            right=Side(style='thin'),
                            top=Side(style='thin'), 
                            bottom=Side(style='thin')
                        )
                
                # Auto-fit columns
                for column in ws.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = min(max_length + 2, 30)
                    ws.column_dimensions[column_letter].width = adjusted_width
                
                # Save to BytesIO
                wb.save(output)
                
                st.download_button(
                    label="📥 Download Excel (Formatted)",
                    data=output.getvalue(),
                    file_name="processed_sales_data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            
            with col2:
                csv_data = final_result.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="📥 Download CSV",
                    data=csv_data,
                    file_name="processed_sales_data.csv",
                    mime="text/csv",
                    use_container_width=True
                )
                    
    except Exception as e:
        st.error(f"Error: {str(e)}")
