import pandas as pd
from datetime import time
import streamlit as st
from io import BytesIO

st.set_page_config(page_title="Bill-wise Item Sales", layout="wide")

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
            
            # Display results
            st.subheader("Processed Results")
            st.dataframe(final_result, use_container_width=True)
            
            # Summary
            st.subheader("Summary")
            col1, col2, col3, col4, col5 = st.columns(5)
            
            with col1:
                st.metric("Total Items", grand_total_all)
            with col2:
                st.metric("Breakfast", grand_total_breakfast)
            with col3:
                st.metric("Lunch", grand_total_lunch)
            with col4:
                st.metric("Snacks", grand_total_snacks)
            with col5:
                st.metric("Late Night", grand_total_latenight)
            
            # Download buttons
            st.subheader("Download")
            col1, col2 = st.columns(2)
            
            with col1:
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    final_result.to_excel(writer, index=False, sheet_name='Sales Data')
                
                st.download_button(
                    label="Download Excel",
                    data=output.getvalue(),
                    file_name="processed_sales_data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            
            with col2:
                csv_data = final_result.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="Download CSV",
                    data=csv_data,
                    file_name="processed_sales_data.csv",
                    mime="text/csv"
                )
                    
    except Exception as e:
        st.error(f"Error: {str(e)}")
