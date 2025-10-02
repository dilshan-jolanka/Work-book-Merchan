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

def find_booking_forms(df):
    """Find all booking forms in the Excel sheet"""
    booking_forms = []
    
    # Look for "Booking Form" text to identify each form
    for row in range(df.shape[0]):
        for col in range(df.shape[1]):
            cell_value = str(df.iloc[row, col]).strip().lower() if pd.notna(df.iloc[row, col]) else ""
            if 'booking form' in cell_value:
                booking_forms.append({'start_row': row, 'start_col': col})
    
    return booking_forms

def extract_single_form_data(df, form_start_row, form_start_col):
    """Extract data from a single booking form starting at given position"""
    base_data = {}
    lot_data = []
    
    # Define the search area for this form (typically 50 rows down from start)
    search_end_row = min(form_start_row + 50, df.shape[0])
    
    # Look for key fields within this form's area
    field_patterns = {
        'Description': ['description', 'desc'],
        'Look': ['look'],
        'Reference': ['ref'],
        'Original Reference': ['original ref'],
        'Supplier Reference': ['supplier ref'],
        'Color': ['color', 'colour'],
        'Total Units': ['uk total unit buy', 'total unit'],
        'VCP': ['vcp'],
        'Factory': ['factory name', 'factory'],
        'Booking Form Delivery': ['booking form delivery', 'booking delivery'],
        'Confirmed Delivery': ['confirmed delivery', 'confirm delivery'],
        'Ship Date': ['ship', 'shipping'],
        'Warehouse Date': ['whs', 'warehouse'],
    }
    
    # Extract base information for this form
    for field, patterns in field_patterns.items():
        for pattern in patterns:
            found = False
            for row in range(form_start_row, search_end_row):
                for col in range(max(0, form_start_col - 2), min(form_start_col + 8, df.shape[1])):
                    cell_value = str(df.iloc[row, col]).strip().lower() if pd.notna(df.iloc[row, col]) else ""
                    if pattern in cell_value:
                        # Look for the actual value in adjacent cells
                        for offset in [1, 2, 3]:
                            if col + offset < df.shape[1]:
                                value_cell = df.iloc[row, col + offset]
                                if pd.notna(value_cell) and str(value_cell).strip() and str(value_cell).strip() != '#N/A':
                                    base_data[field] = str(value_cell).strip()
                                    found = True
                                    break
                        if found:
                            break
                if found:
                    break
            if found:
                break
    
    # Process and format delivery dates
    for date_field in ['Booking Form Delivery', 'Confirmed Delivery', 'Ship Date', 'Warehouse Date']:
        if date_field in base_data:
            raw_date = base_data[date_field]
            formatted_date = format_date(raw_date)
            if formatted_date:
                base_data[f'{date_field}_Formatted'] = formatted_date
    
    # Check if this form has any valid data (not empty/N/A)
    if not base_data or all(val in ['#N/A', '', 'N/A'] for val in base_data.values()):
        return None, []  # Skip empty forms
    
    return base_data, lot_data

def extract_multi_lot_data(df):
    """Extract data from multiple booking forms in the same Excel sheet"""
    all_base_data = []
    all_lot_data = []
    
    # Find all booking forms in the sheet
    booking_forms = find_booking_forms(df)
    
    if not booking_forms:
        # Fallback: treat entire sheet as single form
        booking_forms = [{'start_row': 0, 'start_col': 0}]
    
    # Process each booking form
    for i, form_info in enumerate(booking_forms):
        form_base_data, form_lot_data = extract_single_form_data(
            df, form_info['start_row'], form_info['start_col']
        )
        
        if form_base_data:  # Only add if form has valid data
            # Add form identifier
            form_base_data['Form_Number'] = i + 1
            all_base_data.append(form_base_data)
            
            # If no lot data, create at least one entry from base data
            if not form_lot_data:
                lot_entry = dict(form_base_data)
                lot_entry['Lot Number'] = 1
                all_lot_data.append(lot_entry)
            else:
                all_lot_data.extend(form_lot_data)
    
    # Return combined data from all forms
    return all_base_data, all_lot_data

def process_form_data(base_data_list):
    """Process extracted data from multiple forms"""
    processed_data = []
    
    for base_data in base_data_list:
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
        
        processed_data.append(base_data)
    
    return processed_data

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

def create_order_details_output_multi_form(base_data_list):
    """Create order details output sheet with one row per booking form"""
    
    # Set up dataframe with column headers
    order_details_cols = [
        'FORM_NO', 'IMAGE', 'SUPPLIER REFERENCE', 'DESCRIPTION', 'COLOUR', 'UNITS',
        'BOOKING FORM DELIVERY', 'CONFIRMED DELIVERY', 'VCP', 'FACTORY',
        'FABRIC COMP', 'SUSTAINABLE MESSAGE', 'COST', 'REMARKS'
    ]
    
    if not base_data_list:
        # Return empty dataframe with headers only
        return pd.DataFrame(columns=order_details_cols)
    
    # Create one row per booking form
    order_rows = []
    for i, base_data in enumerate(base_data_list, 1):
        # Skip forms with N/A or empty data
        if (not base_data or 
            base_data.get('Description', '') in ['#N/A', 'N/A', ''] or
            all(val in ['#N/A', 'N/A', ''] for val in base_data.values() if val)):
            continue
            
        # Get delivery dates with fallback logic
        booking_delivery = (base_data.get('Booking Form Delivery_Formatted') or 
                          base_data.get('Ship Date_Formatted') or 
                          base_data.get('Booking Form Delivery') or 
                          base_data.get('Ship Date') or '')
        
        confirmed_delivery = (base_data.get('Confirmed Delivery_Formatted') or 
                            base_data.get('Warehouse Date_Formatted') or 
                            base_data.get('Confirmed Delivery') or 
                            base_data.get('Warehouse Date') or 
                            booking_delivery)  # Use booking delivery as fallback
        
        row = {
            'FORM_NO': i,
            'IMAGE': '',
            'SUPPLIER REFERENCE': base_data.get('Reference', '').upper() if base_data.get('Reference') else '',
            'DESCRIPTION': base_data.get('Description', ''),
            'COLOUR': base_data.get('Color', 'TBC'),
            'UNITS': base_data.get('Total Units', ''),
            'BOOKING FORM DELIVERY': booking_delivery,
            'CONFIRMED DELIVERY': confirmed_delivery,
            'VCP': base_data.get('VCP', ''),
            'FACTORY': base_data.get('Factory', '') + " - " + base_data.get('Factory ID', '') if base_data.get('Factory ID', '') else base_data.get('Factory', ''),
            'FABRIC COMP': '',  # Blank value as requested
            'SUSTAINABLE MESSAGE': '',  # Blank value as requested
            'COST': '',  # Blank value as requested
            'REMARKS': f"Form {base_data.get('Form_Number', i)}"
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
            
            # Extract data from multiple booking forms
            base_data_list, lot_data = extract_multi_lot_data(df)
            
            if base_data_list:
                # Process the extracted data
                processed_base_data = process_form_data(base_data_list)
                
                st.markdown('<div class="info-message">', unsafe_allow_html=True)
                st.write(f"✅ Successfully extracted data from {len(processed_base_data)} booking form(s).")
                if lot_data:
                    st.write(f"✅ Total lots found: {len(lot_data)}")
                st.markdown('</div>', unsafe_allow_html=True)
                
                # Display extracted forms information
                st.markdown("### Extracted Booking Forms")
                
                for i, base_data in enumerate(processed_base_data, 1):
                    with st.expander(f"📋 Booking Form {i} - {base_data.get('Reference', 'No Reference')}"):
                        col1, col2 = st.columns(2)
                        with col1:
                            st.write("**Product Information:**")
                            st.write(f"Description: {base_data.get('Description', 'N/A')}")
                            st.write(f"Reference: {base_data.get('Reference', 'N/A')}")
                            st.write(f"Look: {base_data.get('Look', 'N/A')}")
                            st.write(f"Color: {base_data.get('Color', 'N/A')}")
                        with col2:
                            st.write("**Business Information:**")
                            st.write(f"Factory: {base_data.get('Factory', 'N/A')}")
                            st.write(f"Supplier Ref: {base_data.get('Supplier Reference', 'N/A')}")
                            st.write(f"Total Units: {base_data.get('Total Units', 'N/A')}")
                            st.write(f"VCP: {base_data.get('VCP', 'N/A')}")
                        
                        # Show delivery dates in a separate section
                        st.write("**📅 Delivery Information:**")
                        delivery_col1, delivery_col2 = st.columns(2)
                        with delivery_col1:
                            booking_delivery = (base_data.get('Booking Form Delivery_Formatted') or 
                                              base_data.get('Booking Form Delivery') or 
                                              base_data.get('Ship Date_Formatted') or 
                                              base_data.get('Ship Date') or 'N/A')
                            st.write(f"🚢 Booking Form Delivery: {booking_delivery}")
                        with delivery_col2:
                            confirmed_delivery = (base_data.get('Confirmed Delivery_Formatted') or 
                                                base_data.get('Confirmed Delivery') or 
                                                base_data.get('Warehouse Date_Formatted') or 
                                                base_data.get('Warehouse Date') or 'N/A')
                            st.write(f"✅ Confirmed Delivery: {confirmed_delivery}")
                
                # Show debug information
                with st.expander("🔍 Debug Information"):
                    st.write("**Raw Data Preview:**")
                    st.dataframe(df.head(10))
                    st.write("**Forms Found:**", len(processed_base_data))
                    for i, data in enumerate(processed_base_data):
                        st.write(f"Form {i+1} fields:", list(data.keys()))
                
                # Generate output sheet (one row per form)
                order_df = create_order_details_output_multi_form(processed_base_data)
                
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
                    
                    # Generate filename from first form or use generic name
                    first_ref = processed_base_data[0].get('Reference', 'multi_form') if processed_base_data else 'multi_form'
                    filename = f"order_processing_{first_ref.upper()}_forms.xlsx"
                    
                    st.download_button(
                        label="📥 Download Multi-Form Order Processing Excel",
                        data=output,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="download_button"
                    )
            else:
                st.warning("No valid booking forms found in the Excel sheet.")
                st.info("💡 **Tips for better results:**")
                st.write("- Ensure your Excel sheet contains 'Booking Form' headers")
                st.write("- Make sure forms have valid data (not #N/A)")
                st.write("- Check that key fields like Description, Reference are filled")
                
                # Show what was found for debugging
                with st.expander("🔍 Raw data preview"):
                    st.dataframe(df.head(20))
                
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
