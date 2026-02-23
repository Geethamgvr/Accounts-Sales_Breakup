import pandas as pd
from datetime import time
import streamlit as st
from io import BytesIO

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
    /* Item rows - alternating subtle colors per captain */
    .captain-1-items {
        background-color: #fff2e6 !important;  /* Light orange for captain 1 items */
    }
    .captain-2-items {
        background-color: #e6f0ff !important;  /* Light blue for captain 2 items */
    }
    .captain-3-items {
        background-color: #e6ffe6 !important;  /* Light green for captain 3 items */
    }
    .captain-4-items {
        background-color: #ffe6f0 !important;  /* Light pink for captain 4 items */
    }
    .captain-5-items {
        background-color: #f0e6ff !important;  /* Light purple for captain 5 items */
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
            
            # Captain color mapping
            captain_colors = {}
            for i, captain in enumerate(unique_captains):
                captain_colors[captain] = f'captain-{(i % 5) + 1}-items'
            
            for captain in unique_captains:
                captain_data = file[file['Captain Name'] == captain]
                captain_items = sorted(captain_data['Item Name'].unique())
                
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
                    
                    # Get values and add to grand totals
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
                    
                    row = {
                        'Captain Name': captain_display,
                        'Item Name': item,
                        'Breakfast': breakfast,
                        'Lunch': lunch,
                        'Snacks': snacks,
                        'Late Night': late_night,
                        'Total': total
                    }
                    all_results.append(row)
            
            # Add grand total row at the bottom
            grand_total_row = {
                'Captain Name': 'GRAND TOTAL',
                'Item Name': '--- ALL ITEMS TOTAL ---',
                'Breakfast': grand_total_breakfast,
                'Lunch': grand_total_lunch,
                'Snacks': grand_total_snacks,
                'Late Night': grand_total_latenight,
                'Total': grand_total_all
            }
            all_results.append(grand_total_row)
            
            # Final dataframe
            final_result = pd.DataFrame(all_results)
            
            # Display results with colors and auto-fit
            st.subheader("📋 Processed Results")
            
            # Create color mapping function
            def color_rows(row):
                if row['Captain Name'] == 'GRAND TOTAL':
                    return ['background-color: #d4edda; font-weight: bold'] * len(row)
                elif row['Captain Name'] != '' and row['Captain Name'] != 'GRAND TOTAL':
                    # This is a captain header row
                    return ['background-color: #e6f3ff; font-weight: 500'] * len(row)
                else:
                    # This is an item row - find which captain it belongs to
                    # Find the captain for this item by looking at previous rows
                    for i in range(len(all_results)):
                        if all_results[i]['Item Name'] == row['Item Name']:
                            # Find the captain for this item
                            for j in range(i-1, -1, -1):
                                if all_results[j]['Captain Name'] not in ['', 'GRAND TOTAL']:
                                    captain = all_results[j]['Captain Name']
                                    color_class = captain_colors.get(captain, '')
                                    if color_class == 'captain-1-items':
                                        return ['background-color: #fff2e6'] * len(row)
                                    elif color_class == 'captain-2-items':
                                        return ['background-color: #e6f0ff'] * len(row)
                                    elif color_class == 'captain-3-items':
                                        return ['background-color: #e6ffe6'] * len(row)
                                    elif color_class == 'captain-4-items':
                                        return ['background-color: #ffe6f0'] * len(row)
                                    elif color_class == 'captain-5-items':
                                        return ['background-color: #f0e6ff'] * len(row)
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
            
            # Download buttons
            st.subheader("💾 Download")
            col1, col2 = st.columns(2)
            
            with col1:
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    final_result.to_excel(writer, index=False, sheet_name='Sales Data')
                
                st.download_button(
                    label="📥 Download Excel",
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
