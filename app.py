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
        
        # Show preview
        with st.expander("Preview Uploaded Data"):
            st.dataframe(df.head())
        
        if st.button("Process Data"):
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
                
                st.write(f"Rows after filtering: {len(file)}")
                
                # Create report
                unique_captains = sorted(file['Captain Name'].unique())
                all_results = []
                
                for captain in unique_captains:
                    captain_data = file[file['Captain Name'] == captain]
                    captain_items = sorted(captain_data['Item Name'].unique())
                    
                    # Captain header
                    all_results.append({
                        'Captain Name': captain,
                        'Item Name': '_____ITEM NAME_____',
                        'Breakfast': '',
                        'Lunch': '',
                        'Snacks': '',
                        'Late Night': '',
                        'Total': ''
                    })
                    
                    # Items
                    for item in captain_items:
                        item_data = captain_data[captain_data['Item Name'] == item]
                        pivot = pd.pivot_table(
                            item_data,
                            values='Quantity',
                            columns='Time Category',
                            aggfunc='sum',
                            fill_value=0
                        )
                        
                        row = {
                            'Captain Name': '',
                            'Item Name': item,
                            'Breakfast': pivot.get('Breakfast', 0),
                            'Lunch': pivot.get('Lunch', 0),
                            'Snacks': pivot.get('Snacks', 0),
                            'Late Night': pivot.get('Late Night', 0)
                        }
                        row['Total'] = (row['Breakfast'] + row['Lunch'] + 
                                      row['Snacks'] + row['Late Night'])
                        all_results.append(row)
                    
                    # Captain total
                    captain_total = captain_data.groupby('Time Category')['Quantity'].sum()
                    total_row = {
                        'Captain Name': '',
                        'Item Name': '--- CAPTAIN TOTAL ---',
                        'Breakfast': captain_total.get('Breakfast', 0),
                        'Lunch': captain_total.get('Lunch', 0),
                        'Snacks': captain_total.get('Snacks', 0),
                        'Late Night': captain_total.get('Late Night', 0)
                    }
                    total_row['Total'] = (total_row['Breakfast'] + total_row['Lunch'] + 
                                        total_row['Snacks'] + total_row['Late Night'])
                    all_results.append(total_row)
                    
                    # Empty row
                    all_results.append({
                        'Captain Name': '',
                        'Item Name': '',
                        'Breakfast': '',
                        'Lunch': '',
                        'Snacks': '',
                        'Late Night': '',
                        'Total': ''
                    })
                
                # Final dataframe
                final_result = pd.DataFrame(all_results)
                
                # Display results
                st.subheader("Processed Results")
                
                # Format for display
                display_df = final_result.copy()
                for col in ['Breakfast', 'Lunch', 'Snacks', 'Late Night', 'Total']:
                    display_df[col] = pd.to_numeric(display_df[col], errors='coerce').fillna(0).astype(int)
                
                st.dataframe(display_df, use_container_width=True)
                
                # Summary
                st.subheader("Summary")
                numeric_rows = display_df[pd.to_numeric(display_df['Total'], errors='coerce').notna()]
                
                col1, col2, col3, col4, col5 = st.columns(5)
                with col1:
                    st.metric("Total Items", int(numeric_rows[numeric_rows['Item Name'] != '--- CAPTAIN TOTAL ---']['Total'].sum()))
                with col2:
                    st.metric("Breakfast", int(numeric_rows['Breakfast'].sum()))
                with col3:
                    st.metric("Lunch", int(numeric_rows['Lunch'].sum()))
                with col4:
                    st.metric("Snacks", int(numeric_rows['Snacks'].sum()))
                with col5:
                    st.metric("Late Night", int(numeric_rows['Late Night'].sum()))
                
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
