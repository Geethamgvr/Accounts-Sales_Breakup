import pandas as pd
from datetime import time
import streamlit as st
from io import BytesIO

def process_data(df):
    # Filter for Dine-In orders and remove 'without_captain'
    file1 = df.iloc[:-1] if len(df) > 0 else df
    file2 = file1[file1['Order Type'] == 'Dine-In'].copy()
    file = file2[['Item Name', 'Order Type', 'Quantity', 'Captain Name','Billed Time']]
    file = file[file['Captain Name'].str.lower() != 'without_captain'].copy()
    
    # Convert Billed Time to datetime
    file['Billed Time'] = pd.to_datetime(file['Billed Time'])
    
    # Function to categorize time of day
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
    
    # Add time category column
    file['Time Category'] = file['Billed Time'].apply(get_time_category)
    
    # Function to change the name
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
    
    # Apply to the 'Item Name' column
    file['Item Name'] = file['Item Name'].apply(Grouped_items)
    
    # Remove rows where Item Name is None (unmatched items)
    file = file.dropna(subset=['Item Name']).copy()
    
    st.write(f"Rows after filtering unmatched items: {len(file)}")
    
    # Get unique captains
    unique_captains = sorted(file['Captain Name'].unique())
    
    # Create a list to store results
    all_results = []
    
    for captain in unique_captains:
        # Filter data for current captain
        captain_data = file[file['Captain Name'] == captain]
        
        # Get unique items for this captain
        captain_items = sorted(captain_data['Item Name'].unique())
        
        # Add captain name as a header row (only once per captain)
        all_results.append({
            'Captain Name': captain,
            'Item Name': '',  # Empty for captain header
            'Breakfast': '',
            'Lunch': '',
            'Snacks': '',
            'Late Night': '',
            'Total': ''
        })
        
        # Process each item for this captain
        for item in captain_items:
            # Filter data for current captain and item
            item_data = captain_data[captain_data['Item Name'] == item]
            
            # Create pivot for this captain-item combination
            pivot = pd.pivot_table(
                item_data,
                values='Quantity',
                columns='Time Category',
                aggfunc='sum',
                fill_value=0
            )
            
            # Create a row for this item (empty captain name)
            row_data = {
                'Captain Name': '',  # Empty to avoid repeating captain name
                'Item Name': item,
                'Breakfast': pivot.get('Breakfast', 0),
                'Lunch': pivot.get('Lunch', 0),
                'Snacks': pivot.get('Snacks', 0),
                'Late Night': pivot.get('Late Night', 0)
            }
            
            # Calculate total
            row_data['Total'] = (row_data['Breakfast'] + row_data['Lunch'] + 
                                row_data['Snacks'] + row_data['Late Night'])
            
            all_results.append(row_data)
        
        # Add captain total row
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
        
        # Add empty row between captains
        all_results.append({
            'Captain Name': '',
            'Item Name': '',
            'Breakfast': '',
            'Lunch': '',
            'Snacks': '',
            'Late Night': '',
            'Total': ''
        })
    
    # Create final dataframe
    final_result = pd.DataFrame(all_results)
    return final_result

# Streamlit app
st.set_page_config(page_title="Bill-wise Item Sales Processor", layout="wide")

st.title("Bill-wise Item Sales Processor")
st.write("Upload your CSV file to process the sales data")

# File upload
uploaded_file = st.file_uploader("Choose a CSV file", type="csv")

if uploaded_file is not None:
    try:
        # Read the uploaded file
        df = pd.read_csv(uploaded_file, skiprows=5)
        
        st.success("File uploaded successfully!")
        
        # Display sample of uploaded data
        st.subheader("Sample of Uploaded Data")
        st.dataframe(df.head())
        
        # Process button
        if st.button("Process Data"):
            with st.spinner("Processing data..."):
                # Process the data
                result_df = process_data(df)
                
                # Display results
                st.subheader("Processed Results")
                
                # Format numerical columns for display
                display_df = result_df.copy()
                numeric_cols = ['Breakfast', 'Lunch', 'Snacks', 'Late Night', 'Total']
                for col in numeric_cols:
                    display_df[col] = pd.to_numeric(display_df[col], errors='coerce').fillna(0).astype(int)
                
                st.dataframe(display_df, use_container_width=True)
                
                # Create download button for processed data
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    result_df.to_excel(writer, index=False, sheet_name='Sales Data')
                
                st.download_button(
                    label="Download Processed Data (Excel)",
                    data=output.getvalue(),
                    file_name="processed_sales_data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                # Also provide CSV download option
                csv_data = result_df.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="Download Processed Data (CSV)",
                    data=csv_data,
                    file_name="processed_sales_data.csv",
                    mime="text/csv"
                )
                
                # Show summary statistics
                st.subheader("Summary Statistics")
                
                # Calculate totals only from numeric rows
                numeric_rows = display_df[pd.to_numeric(display_df['Total'], errors='coerce').notna()]
                
                col1, col2, col3, col4, col5 = st.columns(5)
                
                with col1:
                    st.metric("Total Items Sold", int(numeric_rows[numeric_rows['Item Name'] != '--- CAPTAIN TOTAL ---']['Total'].sum()))
                with col2:
                    st.metric("Breakfast Total", int(numeric_rows['Breakfast'].sum()))
                with col3:
                    st.metric("Lunch Total", int(numeric_rows['Lunch'].sum()))
                with col4:
                    st.metric("Snacks Total", int(numeric_rows['Snacks'].sum()))
                with col5:
                    st.metric("Late Night Total", int(numeric_rows['Late Night'].sum()))
                
    except Exception as e:
        st.error(f"Error processing file: {str(e)}")
else:
    st.info("Please upload a CSV file to begin processing.")

# Instructions
with st.expander("Instructions"):
    st.write("""
    1. Upload your CSV file (should be in the same format as Bill-wise Item Sales export)
    2. Click 'Process Data' to analyze the sales
    3. View the processed results in the table
    4. Download the processed data as Excel or CSV file
    5. Summary statistics are shown at the bottom
    
    **Note:** The file should have columns: Item Name, Order Type, Quantity, Captain Name, Billed Time
    """)
