import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import re
from datetime import datetime, timedelta
import os
from dotenv import load_dotenv

# Load env variables
load_dotenv()

def format_date(date_str):
    """Convert date to DD-MMM format"""
    try:
        # Handle formats like "19 Jul '25" or "2025-07-19 00:00:00"
        if isinstance(date_str, str):
            if "'" in date_str:  # Format like "19 Jul '25"
                parts = date_str.split()
                if len(parts) >= 2:
                    return f"{parts[0]}-{parts[1]}"
            elif "00:00:00" in date_str:  # Format like "2025-07-19 00:00:00"
                date_part = date_str.split()[0]
                if "-" in date_part:
                    year, month, day = date_part.split('-')
                    month_names = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", 
                                  "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
                    return f"{int(day)}-{month_names[int(month)-1]}"
        elif isinstance(date_str, datetime):
            return f"{date_str.day}-{date_str.strftime('%b')}"
    except Exception as e:
        pass
    return ""

def extract_multi_lot_data(df):
    """Extract data for multiple lots from the booking form"""
    base_data = {}
    lot_data = []
    
    # Extract basic information that applies to all lots
    base_mappings = {
        'Description': (20, 3),       # Row 20, Column 3
        'Reference': (22, 3),         # Row 22, Column 3
        'Original Reference': (23, 3), # Row 23, Column 3
        'Supplier Reference': (25, 3), # Row 25, Column 3
        'Color': (26, 3),             # Row 26, Column 3
        'Total Units': (27, 3),       # Row 27, Column 3 - Total units across all lots
        'VCP': (28, 3),               # Row 28, Column 3
        'Factory': (14, 3),           # Row 14, Column 3 (Factory Name)
    }
    
    # Extract base data
    for field, (row, col) in base_mappings.items():
        if row < df.shape[0] and col < df.shape[1]:
            cell_value = df.iloc[row, col]
            if pd.notna(cell_value):
                if isinstance(cell_value, (int, float)):
                    base_data[field] = str(int(cell_value)) if cell_value.is_integer() else str(cell_value)
                else:
                    base_data[field] = str(cell_value).strip()
    
    # Process factory name
    if 'Factory' in base_data:
        factory_value = base_data['Factory']
        if '[' in factory_value:
            parts = factory_value.split('[')
            if len(parts) >= 2:
                base_data['Factory'] = parts[0].strip()
                base_data['Factory ID'] = parts[1].replace(']', '').strip()
    
    # Process color
    if 'Color' in base_data:
        color_value = base_data['Color']
        if '[' in color_value:
            base_data['Color'] = color_value.split('[')[0].strip()
            color_code = re.search(r'\[(.*?)\]', color_value)
            if color_code:
                base_data['Color Code'] = color_code.group(1)
    
    # Find the ship and warehouse date rows
    ship_row = None
    whs_row = None
    for i in range(15, 25):  # Look around rows 20-21 where these often appear
        if i < df.shape[0]:
            for j in range(7, 8):  # Look in column 7 for the labels
                if j < df.shape[1]:
                    cell_value = str(df.iloc[i, j]).strip() if pd.notna(df.iloc[i, j]) else ""
                    if 'Ship' in cell_value:
                        ship_row = i
                    elif 'Whs' in cell_value:
                        whs_row = i
    
    # Find the units row
    units_row = None
    for i in range(18, 25):
        if i < df.shape[0]:
            cell_value = str(df.iloc[i, 7]).strip() if pd.notna(df.iloc[i, 7]) else ""
            if 'Units' in cell_value:
                units_row = i
                break
    
    # Determine how many lots are in the form by checking for non-empty units
    if units_row is not None:
        lot_count = 0
        for col in range(9, 17):  # Check columns 9-16 for lot data
            if col < df.shape[1]:
                units_value = df.iloc[units_row, col]
                if pd.notna(units_value) and units_value != 0:
                    lot_count += 1
                    
                    # Create a lot entry
                    lot_entry = dict(base_data)  # Copy base data
                    
                    # Add lot-specific data
                    lot_entry['Lot Number'] = lot_count
                    lot_entry['Units'] = str(int(units_value)) if isinstance(units_value, (int, float)) else str(units_value).strip()
                    
                    # Add ship date if available
                    if ship_row is not None and col < df.shape[1]:
                        ship_date = df.iloc[ship_row, col]
                        if pd.notna(ship_date):
                            lot_entry['Ship Date'] = str(ship_date)
                            lot_entry['Ship Date Formatted'] = format_date(ship_date)
                            
                            # Calculate Ex FTY date (12 days before Ship Date)
                            try:
                                date_str = str(ship_date)
                                if ' ' in date_str and "00:00:00" in date_str:  # Format like "2025-07-19 00:00:00"
                                    date_parts = date_str.split(' ')[0].split('-')
                                    if len(date_parts) == 3:
                                        year, month, day = date_parts
                                        ship_datetime = datetime(int(year), int(month), int(day))
                                        ex_fty_date = ship_datetime - timedelta(days=12)
                                        lot_entry['Ex FTY'] = f"{ex_fty_date.day}-{ex_fty_date.strftime('%b')}"
                                elif "'" in date_str:  # Format like "19 Jul '25"
                                    date_parts = date_str.split()
                                    if len(date_parts) >= 3:
                                        day = int(date_parts[0])
                                        month_str = date_parts[1]
                                        year_str = date_parts[2].replace("'", "20")
                                        month_map = {"Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4, "May": 5, "Jun": 6, 
                                                    "Jul": 7, "Aug": 8, "Sep": 9, "Oct": 10, "Nov": 11, "Dec": 12}
                                        month = month_map.get(month_str, 1)
                                        year = int(year_str)
                                        ship_datetime = datetime(year, month, day)
                                        ex_fty_date = ship_datetime - timedelta(days=12)
                                        lot_entry['Ex FTY'] = f"{ex_fty_date.day}-{ex_fty_date.strftime('%b')}"
                            except Exception as e:
                                pass
                    
                    # Add warehouse date if available
                    if whs_row is not None and col < df.shape[1]:
                        whs_date = df.iloc[whs_row, col]
                        if pd.notna(whs_date):
                            lot_entry['Warehouse Date'] = str(whs_date)
                            lot_entry['Warehouse Date Formatted'] = format_date(whs_date)
                    
                    lot_data.append(lot_entry)
    
    return base_data, lot_data

def create_order_details_output(base_data, lot_data):
    """Create order details output sheet with one row per lot"""
    
    # Set up empty dataframe with column headers only
    order_details_cols = [
        'IMAGE', 'SUPPLIER REFERENCE', 'DESCRIPTION', 'COLOUR', 'UNITS',
        'BOOKING FORM DELIVERY', 'CONFIRMED DELIVERY', 'VCP', 'FACTORY',
        'FABRIC COMP', 'SUSTAINABLE MESSAGE', 'COST', 'REMARKS'
    ]
    
    if not lot_data:
        # Return empty dataframe with headers only
        return pd.DataFrame(columns=order_details_cols)
    
    # Create Sheet 1 - one row per lot
    order_rows = []
    for lot in lot_data:
        row = {
            'IMAGE': '',
            'SUPPLIER REFERENCE': lot.get('Reference', '').upper(),
            'DESCRIPTION': lot.get('Description', ''),
            'COLOUR': lot.get('Color', 'TBC'),
            'UNITS': lot.get('Units', ''),
            'BOOKING FORM DELIVERY': lot.get('Ship Date Formatted', ''),
            'CONFIRMED DELIVERY': lot.get('Ship Date Formatted', ''),  # Same as booking form delivery
            'VCP': lot.get('VCP', ''),
            'FACTORY': lot.get('Factory', '') + " - " + lot.get('Factory ID', '') if lot.get('Factory ID', '') else lot.get('Factory', ''),
            'FABRIC COMP': '',  # Blank value as requested
            'SUSTAINABLE MESSAGE': '',  # Blank value as requested
            'COST': '',  # Blank value as requested
            'REMARKS': ''
        }
        order_rows.append(row)
    
    return pd.DataFrame(order_rows)

def allow_manual_edits(order_df):
    """Allow users to manually edit the generated data before final output"""
    st.markdown("### Edit Order Details")
    st.write("Make any needed changes to the data before generating the final Excel file.")
    
    edited_order = st.data_editor(
        order_df,
        num_rows="fixed",
        use_container_width=True,
        hide_index=True
    )
    
    return edited_order

def main():
    st.set_page_config(page_title="Booking Form Processor", layout="wide")
    
    st.markdown("""
        <style>
            .main-title { font-size: 40px; color: #4F8BF9; text-align: center; margin-bottom: 20px; }
            .sub-title { font-size: 20px; color: #4F8BF9; margin-top: 20px; }
            .success-message {
                background: #ccffcc;
                border: 1px solid #00cc00;
                border-radius: 5px;
                padding: 15px;
                margin: 20px 0;
                font-weight: bold;
                color: #006600;
                text-align: center;
            }
            .info-message {
                background: #e6f2ff;
                border: 1px solid #0066cc;
                border-radius: 5px;
                padding: 15px;
                margin: 20px 0;
                font-weight: bold;
                color: #0044cc;
                text-align: center;
            }
            .preview-container {
                margin: 20px 0;
                border: 1px solid #ddd;
                border-radius: 5px;
                padding: 15px;
                background-color: #f9f9f9;
            }
        </style>
    """, unsafe_allow_html=True)

    st.markdown('<h1 class="main-title">Booking Form Processor</h1>', unsafe_allow_html=True)
    
    st.write("Upload an Excel booking form and generate order details with one line per lot.")
    
    # Use the specific date and time provided by the user
    current_time = "2025-04-05 05:00:44"
    current_user = "dilshan-jolanka"
    
    st.sidebar.markdown("### System Information")
    st.sidebar.write(f"**Date & Time (UTC):** {current_time}")
    st.sidebar.write(f"**Current User:** {current_user}")

    uploaded_excel = st.file_uploader("Upload Excel booking form", type=["xlsx", "xls"])
    
    if uploaded_excel:
        st.success(f"File uploaded: {uploaded_excel.name}")
        
        try:
            # Choose the appropriate engine based on file extension
            file_extension = uploaded_excel.name.split('.')[-1].lower()
            
            if file_extension == 'xls':
                engine = 'xlrd'
            else:  # xlsx or other formats
                engine = 'openpyxl'
                
            # Read Excel file - using the appropriate engine
            df = pd.read_excel(uploaded_excel, header=None, engine=engine)
            
            # Extract data for multiple lots
            base_data, lot_data = extract_multi_lot_data(df)
            
            if lot_data:
                st.markdown('<div class="info-message">', unsafe_allow_html=True)
                st.write(f"✅ Successfully extracted {len(lot_data)} lots from the booking form.")
                st.markdown('</div>', unsafe_allow_html=True)
                
                # Display lot information
                st.markdown("### Extracted Lot Data")
                
                for i, lot in enumerate(lot_data, 1):
                    col1, col2 = st.columns(2)
                    with col1:
                        st.write(f"**Lot {i}:**")
                        st.write(f"Units: {lot.get('Units', 'N/A')}")
                        st.write(f"Ship Date: {lot.get('Ship Date Formatted', 'N/A')}")
                    with col2:
                        st.write(f"Reference: {lot.get('Reference', 'N/A')}")
                        st.write(f"Warehouse Date: {lot.get('Warehouse Date Formatted', 'N/A')}")
                        st.write(f"Ex FTY: {lot.get('Ex FTY', 'N/A')}")
                
                # Generate output sheet (one row per lot)
                order_df = create_order_details_output(base_data, lot_data)
                
                # Allow manual editing
                edited_order = allow_manual_edits(order_df)
                
                # Create Excel output
                if st.button("Generate Order Processing Sheet", type="primary"):
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        edited_order.to_excel(writer, sheet_name='Order Details', index=False)
                        
                        # Add formatting
                        workbook = writer.book
                        header_format = workbook.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 1})
                        
                        # Format the worksheet
                        worksheet = writer.sheets['Order Details']
                        for col_num, value in enumerate(edited_order.columns):
                            worksheet.write(0, col_num, value, header_format)
                            worksheet.set_column(col_num, col_num, max(12, len(value) + 2))
                    
                    output.seek(0)
                    
                    st.markdown('<div class="success-message">', unsafe_allow_html=True)
                    st.markdown("✅ **Order processing sheet has been generated successfully!**")
                    st.markdown('</div>', unsafe_allow_html=True)
                    
                    ref = base_data.get('Reference', 'form').upper()
                    st.download_button(
                        label="Download Order Processing Excel",
                        data=output,
                        file_name=f"order_processing_{ref}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="download_button"
                    )
            else:
                st.warning("No lot information could be extracted from the booking form.")
                
        except Exception as e:
            st.error(f"Error processing Excel file: {str(e)}")
            import traceback
            st.error(traceback.format_exc())
            
            # Add helpful troubleshooting information
            st.markdown("""
            ### Troubleshooting Excel File Issues
            
            If you're having trouble with your Excel file, try these steps:
            
            1. Ensure you have both libraries installed:
               ```
               pip install xlrd>=2.0.1 openpyxl
               ```
            
            2. Make sure your Excel file is properly formatted
            
            3. Try saving your Excel file in a different format (.xls if you have .xlsx or vice versa)
            """)

if __name__ == "__main__":
    main()