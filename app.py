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
    st.markdown('''
        <div class="modern-card">
            <h3 style="color: #00f5ff; text-align: center; margin-bottom: 1rem;">
                ✏️ DATA EDITOR PANEL
            </h3>
            <p style="color: rgba(255,255,255,0.8); text-align: center; margin-bottom: 2rem;">
                Review and modify the extracted data before generating the final Excel file
            </p>
        </div>
    ''', unsafe_allow_html=True)
    
    # Enhanced data editor with styling
    edited_order = st.data_editor(
        order_df,
        num_rows="fixed",
        use_container_width=True,
        hide_index=True,
        column_config={
            "SUPPLIER REFERENCE": st.column_config.TextColumn("🏷️ Supplier Ref", width="medium"),
            "DESCRIPTION": st.column_config.TextColumn("📝 Description", width="large"),
            "COLOUR": st.column_config.TextColumn("🎨 Color", width="small"),
            "UNITS": st.column_config.NumberColumn("📦 Units", width="small"),
            "BOOKING FORM DELIVERY": st.column_config.TextColumn("🚢 Booking Delivery", width="medium"),
            "CONFIRMED DELIVERY": st.column_config.TextColumn("✅ Confirmed Delivery", width="medium"),
            "VCP": st.column_config.TextColumn("💰 VCP", width="small"),
            "FACTORY": st.column_config.TextColumn("🏭 Factory", width="medium"),
        }
    )
    
    return edited_order

def main():
    st.set_page_config(
        page_title="🚀 Jolanka AI Booking Processor", 
        layout="wide",
        page_icon="🚀",
        initial_sidebar_state="expanded"
    )
    
    # Professional Dark Theme CSS styling
    st.markdown("""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;500;600&display=swap');
        
        /* Global dark theme override */
        .stApp {
            background: #0a0a0a !important;
            color: #ffffff !important;
        }
        
        .main {
            background: #0a0a0a !important;
            padding: 1rem 2rem !important;
        }
        
        /* Container styling */
        .block-container {
            background: #0a0a0a !important;
            padding: 1rem 2rem !important;
            max-width: 1400px !important;
        }
        
        /* Override all white backgrounds */
        div[data-testid="stAppViewContainer"] {
            background: #0a0a0a !important;
        }
        
        div[data-testid="stHeader"] {
            background: transparent !important;
        }
        
        section[data-testid="stSidebar"] {
            background: #111111 !important;
        }
        
        /* Content wrapper */
        .block-container {
            background: transparent;
            padding-top: 1rem;
            padding-bottom: 1rem;
        }
        
        /* Hide default Streamlit elements */
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        header {visibility: hidden;}
        
        /* Professional main title */
        .main-title {
            font-family: 'Inter', sans-serif;
            font-size: 2.8rem;
            font-weight: 800;
            color: #ffffff;
            text-align: center;
            margin: 1.5rem 0 0.5rem 0;
            letter-spacing: -0.02em;
            text-shadow: 0 2px 4px rgba(0, 0, 0, 0.3);
        }
        
        @keyframes gradientShift {
            0% { background-position: 0% 50%; }
            50% { background-position: 100% 50%; }
            100% { background-position: 0% 50%; }
        }
        
        /* Professional subtitle */
        .sub-title {
            font-family: 'Inter', sans-serif;
            font-size: 1.1rem;
            color: #94a3b8;
            text-align: center;
            margin-bottom: 2rem;
            font-weight: 400;
            letter-spacing: 0.01em;
        }
        
        /* Professional card styling */
        .modern-card {
            background: #1a1a1a;
            border: 1px solid #2d2d2d;
            border-radius: 12px;
            padding: 1.5rem;
            margin: 1rem 0;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1);
            transition: all 0.3s ease;
        }
        
        .modern-card:hover {
            border-color: #3b82f6;
            box-shadow: 0 8px 24px rgba(0, 0, 0, 0.2);
            transform: translateY(-1px);
        }
        
        /* Professional success message */
        .success-message {
            background: #065f46;
            border: 1px solid #10b981;
            border-radius: 8px;
            padding: 1rem 1.5rem;
            text-align: center;
            font-weight: 500;
            color: #d1fae5;
            font-family: 'Inter', sans-serif;
            margin: 1rem 0;
        }
        
        @keyframes pulse {
            0% { box-shadow: 0 0 20px rgba(0, 255, 128, 0.3); }
            50% { box-shadow: 0 0 30px rgba(0, 255, 128, 0.5); }
            100% { box-shadow: 0 0 20px rgba(0, 255, 128, 0.3); }
        }
        
        /* Professional info message */
        .info-message {
            background: #1e3a8a;
            border: 1px solid #3b82f6;
            border-radius: 8px;
            padding: 1rem 1.5rem;
            text-align: center;
            font-weight: 500;
            color: #dbeafe;
            font-family: 'Inter', sans-serif;
            margin: 1rem 0;
        }
        
        /* Warning message */
        .warning-message {
            background: linear-gradient(45deg, rgba(255, 165, 0, 0.1), rgba(255, 69, 0, 0.1));
            border: 2px solid #ff8c00;
            border-radius: 15px;
            padding: 1.5rem;
            text-align: center;
            font-weight: 600;
            color: #ff8c00;
            box-shadow: 0 0 20px rgba(255, 140, 0, 0.3);
        }
        
        /* Error message */
        .error-message {
            background: linear-gradient(45des, rgba(255, 0, 128, 0.1), rgba(255, 0, 0, 0.1));
            border: 2px solid #ff0080;
            border-radius: 15px;
            padding: 1.5rem;
            text-align: center;
            font-weight: 600;
            color: #ff0080;
            box-shadow: 0 0 20px rgba(255, 0, 128, 0.3);
        }
        
        /* Professional upload area */
        .upload-container {
            background: #1a1a1a;
            border: 2px dashed #374151;
            border-radius: 12px;
            padding: 2rem;
            text-align: center;
            transition: all 0.3s ease;
            margin: 1.5rem 0;
        }
        
        .upload-container:hover {
            border-color: #3b82f6;
            background: #1e1e1e;
            box-shadow: 0 4px 12px rgba(59, 130, 246, 0.1);
        }
        
        /* File uploader styling */
        .stFileUploader {
            background: transparent !important;
        }
        
        .stFileUploader > div {
            background: #1a1a1a !important;
            border: 2px dashed #374151 !important;
            border-radius: 12px !important;
            padding: 2rem !important;
        }
        
        .stFileUploader > div:hover {
            border-color: #3b82f6 !important;
            background: #1e1e1e !important;
        }
        
        .stFileUploader label {
            color: #ffffff !important;
            font-family: 'Inter', sans-serif !important;
            font-weight: 500 !important;
        }
        
        /* Data frames and tables */
        .dataframe {
            background: rgba(255, 255, 255, 0.05) !important;
            color: #ffffff !important;
            border-radius: 15px !important;
        }
        
        .dataframe th {
            background: rgba(0, 245, 255, 0.2) !important;
            color: #00f5ff !important;
            border: 1px solid rgba(0, 245, 255, 0.3) !important;
        }
        
        .dataframe td {
            background: rgba(255, 255, 255, 0.02) !important;
            color: rgba(255, 255, 255, 0.9) !important;
            border: 1px solid rgba(0, 245, 255, 0.1) !important;
        }
        
        /* Expander content background */
        .streamlit-expanderContent {
            background: rgba(0, 0, 0, 0.3) !important;
            border-radius: 0 0 15px 15px;
        }
        
        /* Stats container */
        .stats-container {
            display: flex;
            justify-content: space-around;
            margin: 2rem 0;
        }
        
        .stat-item {
            background: rgba(255, 255, 255, 0.05);
            border: 1px solid rgba(0, 245, 255, 0.3);
            border-radius: 15px;
            padding: 1.5rem;
            text-align: center;
            min-width: 150px;
            backdrop-filter: blur(5px);
        }
        
        .stat-number {
            font-family: 'Orbitron', monospace;
            font-size: 2rem;
            font-weight: 700;
            color: #00f5ff;
            margin-bottom: 0.5rem;
        }
        
        .stat-label {
            font-family: 'Roboto', sans-serif;
            color: rgba(255, 255, 255, 0.8);
            font-size: 0.9rem;
        }
        
        /* Professional button styling */
        .stButton > button {
            background: #3b82f6;
            border: none;
            border-radius: 8px;
            padding: 0.75rem 1.5rem;
            font-family: 'Inter', sans-serif;
            font-weight: 500;
            color: white;
            transition: all 0.2s ease;
            font-size: 0.95rem;
        }
        
        .stButton > button:hover {
            background: #2563eb;
            transform: translateY(-1px);
            box-shadow: 0 4px 12px rgba(59, 130, 246, 0.3);
        }
        
        .stDownloadButton > button {
            background: #10b981;
            border: none;
            border-radius: 8px;
            padding: 0.75rem 1.5rem;
            font-family: 'Inter', sans-serif;
            font-weight: 500;
            color: white;
            transition: all 0.2s ease;
        }
        
        .stDownloadButton > button:hover {
            background: #059669;
            transform: translateY(-1px);
            box-shadow: 0 4px 12px rgba(16, 185, 129, 0.3);
        }
        
        /* Sidebar styling */
        .css-1d391kg, .css-1y4p8pa, .css-17eq0hr {
            background: linear-gradient(180deg, #0a0a1f 0%, #1a1a2e 50%, #2d1b69 100%) !important;
        }
        
        /* Sidebar content */
        .css-1d391kg .element-container {
            background: transparent;
        }
        
        /* Remove white backgrounds from all containers */
        .element-container, .stMarkdown, .stText {
            background: transparent !important;
        }
        
        /* File uploader styling */
        .stFileUploader {
            background: rgba(0, 245, 255, 0.05) !important;
            border-radius: 15px;
            border: 2px dashed rgba(0, 245, 255, 0.3);
            padding: 2rem;
        }
        
        .stFileUploader label {
            color: #00f5ff !important;
            font-weight: 600;
        }
        
        /* Text styling */
        p, span, div {
            color: rgba(255, 255, 255, 0.9) !important;
        }
        
        h1, h2, h3, h4, h5, h6 {
            color: #00f5ff !important;
        }
        
        /* Remove default streamlit styling */
        .css-1v0mbdj, .css-18e3th9, .css-1d391kg {
            background: transparent !important;
        }
        
        /* Data editor styling */
        .stDataFrame {
            border-radius: 15px;
            overflow: hidden;
            box-shadow: 0 8px 32px rgba(0, 245, 255, 0.1);
        }
        
        /* Progress bar */
        .progress-container {
            background: rgba(255, 255, 255, 0.1);
            border-radius: 10px;
            padding: 0.5rem;
            margin: 1rem 0;
        }
        
        .progress-bar {
            background: linear-gradient(90deg, #00f5ff, #8000ff);
            height: 8px;
            border-radius: 4px;
            transition: width 0.3s ease;
        }
        
        /* Expander styling */
        .streamlit-expanderHeader {
            background: rgba(0, 245, 255, 0.1);
            border-radius: 10px;
            border: 1px solid rgba(0, 245, 255, 0.3);
        }
        
        /* Custom text colors */
        .highlight-text {
            color: #00f5ff;
            font-weight: 600;
        }
        
        .accent-text {
            color: #ff0080;
            font-weight: 500;
        }
        
        .success-text {
            color: #00ff80;
            font-weight: 500;
        }
        </style>
    """, unsafe_allow_html=True)

    # Main header with single logo and animation
    st.markdown('''
        <div style="text-align: center; margin: 2rem 0;">
            <img src="https://jolankagroup.com/wp-content/themes/jolanka/assets/images/icons/jolanka-logo-no-text.png" 
                 style="height: 80px; margin-bottom: 1rem; filter: drop-shadow(0 0 20px rgba(0, 245, 255, 0.5));" 
                 alt="Jolanka Logo">
            <div class="main-title">JOLANKA AI PROCESSOR</div>
            <div class="sub-title">Advanced Multi-Form Excel Data Extraction Platform</div>
        </div>
    ''', unsafe_allow_html=True)
    
    # Enhanced sidebar with modern design
    with st.sidebar:
        st.markdown('''
            <div style="text-align: center; margin: 2rem 0;">
                <h3 style="color: #00f5ff; font-family: 'Orbitron', monospace;">CONTROL PANEL</h3>
            </div>
        ''', unsafe_allow_html=True)
        
        # Professional System Information
        st.markdown('''
            <div class="modern-card">
                <h4 style="color: #ffffff; margin-bottom: 1rem; font-family: 'Inter', sans-serif; font-weight: 600;">System Status</h4>
                <div style="font-family: 'Inter', sans-serif; font-size: 0.9rem;">
                    <div style="margin: 0.5rem 0; display: flex; justify-content: space-between;">
                        <span style="color: #94a3b8;">Status:</span> 
                        <span style="color: #10b981;">ACTIVE</span>
                    </div>
                    <div style="margin: 0.5rem 0; display: flex; justify-content: space-between;">
                        <span style="color: #94a3b8;">User:</span> 
                        <span style="color: #3b82f6;">dilshan-jolanka</span>
                    </div>
                    <div style="margin: 0.5rem 0; display: flex; justify-content: space-between;">
                        <span style="color: #94a3b8;">Date:</span> 
                        <span style="color: #ffffff;">2025-10-02</span>
                    </div>
                </div>
            </div>
        ''', unsafe_allow_html=True)
        
        # Features panel
        st.markdown('''
            <div class="modern-card">
                <h4 style="color: #00f5ff; margin-bottom: 1rem;">⚡ FEATURES</h4>
                <div style="color: rgba(255,255,255,0.8); font-size: 0.9rem;">
                    <div style="margin: 0.5rem 0; display: flex; align-items: center; font-size: 0.9rem;">
                        <span style="color: #10b981; margin-right: 0.5rem;">✓</span>
                        <span style="color: #e2e8f0;">Multi-Form Detection</span>
                    </div>
                    <div style="margin: 0.5rem 0; display: flex; align-items: center; font-size: 0.9rem;">
                        <span style="color: #10b981; margin-right: 0.5rem;">✓</span>
                        <span style="color: #e2e8f0;">AI-Powered Extraction</span>
                    </div>
                    <div style="margin: 0.5rem 0; display: flex; align-items: center; font-size: 0.9rem;">
                        <span style="color: #10b981; margin-right: 0.5rem;">✓</span>
                        <span style="color: #e2e8f0;">Delivery Date Processing</span>
                    </div>
                    <div style="margin: 0.5rem 0; display: flex; align-items: center; font-size: 0.9rem;">
                        <span style="color: #10b981; margin-right: 0.5rem;">✓</span>
                        <span style="color: #e2e8f0;">Advanced Data Validation</span>
                    </div>
                    <div style="margin: 0.5rem 0; display: flex; align-items: center; font-size: 0.9rem;">
                        <span style="color: #10b981; margin-right: 0.5rem;">✓</span>
                        <span style="color: #e2e8f0;">Excel Export Ready</span>
                    </div>
                </div>
            </div>
        ''', unsafe_allow_html=True)
    
    # Main content area with enhanced upload section
    st.markdown('''
        <div class="modern-card" style="background: rgba(0, 245, 255, 0.03); border: 2px solid rgba(0, 245, 255, 0.2);">
            <div style="text-align: center;">
                <div style="font-size: 4rem; margin-bottom: 1rem; animation: float 3s ease-in-out infinite;">📁</div>
                <h3 style="color: #00f5ff; margin-bottom: 1rem; font-family: 'Orbitron', monospace;">FILE UPLOAD CENTER</h3>
                <p style="color: rgba(255,255,255,0.8); margin-bottom: 1rem;">
                    Upload your Excel booking forms for intelligent data extraction and processing
                </p>
                <div style="display: flex; justify-content: center; gap: 2rem; margin-top: 1.5rem;">
                    <div style="text-align: center;">
                        <div style="color: #00ff80; font-size: 1.5rem;">✓</div>
                        <small style="color: rgba(255,255,255,0.7);">Multi-Form Detection</small>
                    </div>
                    <div style="text-align: center;">
                        <div style="color: #ff0080; font-size: 1.5rem;">⚡</div>
                        <small style="color: rgba(255,255,255,0.7);">Smart Processing</small>
                    </div>
                    <div style="text-align: center;">
                        <div style="color: #8000ff; font-size: 1.5rem;">📊</div>
                        <small style="color: rgba(255,255,255,0.7);">Auto Export</small>
                    </div>
                </div>
            </div>
        </div>
        
        <style>
        @keyframes float {
            0% { transform: translateY(0px); }
            50% { transform: translateY(-10px); }
            100% { transform: translateY(0px); }
        }
        </style>
    ''', unsafe_allow_html=True)

    # Create a compact main content area
    st.markdown('<div style="margin-top: 1rem;">', unsafe_allow_html=True)
    
    # File upload section
    st.markdown('''
        <div class="modern-card" style="margin-bottom: 1.5rem;">
            <h3 style="color: #ffffff; margin-bottom: 1rem; font-family: 'Inter', sans-serif; font-weight: 600; font-size: 1.2rem;">
                📄 Upload Excel File
            </h3>
            <p style="color: #94a3b8; margin-bottom: 1rem; font-size: 0.9rem;">
                Select your Excel booking form containing multiple forms for processing
            </p>
        </div>
    ''', unsafe_allow_html=True)
    
    uploaded_excel = st.file_uploader(
        "Choose file", 
        type=["xlsx", "xls"],
        help="Supported formats: .xlsx, .xls",
        label_visibility="collapsed"
    )
    
    if uploaded_excel:
        # Professional file upload confirmation
        st.markdown(f'''
            <div class="success-message">
                ✅ <strong>File uploaded successfully:</strong> {uploaded_excel.name} ({uploaded_excel.size:,} bytes)
            </div>
        ''', unsafe_allow_html=True)
        
        # Add a loading progress bar
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        try:
            status_text.text("🔍 Analyzing file format...")
            progress_bar.progress(20)
            # Choose the appropriate engine based on file extension
            file_extension = uploaded_excel.name.split('.')[-1].lower()
            
            if file_extension == 'xls':
                engine = 'xlrd'
            else:  # xlsx or other formats
                engine = 'openpyxl'
            
            status_text.text("📊 Loading Excel data...")
            progress_bar.progress(40)
            
            # Read Excel file - using the appropriate engine
            df = pd.read_excel(uploaded_excel, header=None, engine=engine)
            
            status_text.text("🤖 AI Processing booking forms...")
            progress_bar.progress(60)
            
            # Extract data from multiple booking forms
            base_data_list, lot_data = extract_multi_lot_data(df)
            
            status_text.text("✨ Finalizing data extraction...")
            progress_bar.progress(80)
            
            if base_data_list:
                # Process the extracted data
                processed_base_data = process_form_data(base_data_list)
                
                progress_bar.progress(100)
                status_text.text("🎯 Extraction Complete!")
                
                # Enhanced success display with statistics
                valid_forms = len(processed_base_data)
                total_units = sum(int(form.get('Total Units', '0').replace(',', '') or '0') for form in processed_base_data if form.get('Total Units', '0').replace(',', '').isdigit())
                
                st.markdown(f'''
                    <div class="success-message">
                        <h3>🎉 DATA EXTRACTION SUCCESSFUL</h3>
                        <div class="stats-container">
                            <div class="stat-item">
                                <div class="stat-number">{valid_forms}</div>
                                <div class="stat-label">Forms Processed</div>
                            </div>
                            <div class="stat-item">
                                <div class="stat-number">{total_units:,}</div>
                                <div class="stat-label">Total Units</div>
                            </div>
                            <div class="stat-item">
                                <div class="stat-number">{df.shape[0]}</div>
                                <div class="stat-label">Rows Analyzed</div>
                            </div>
                            <div class="stat-item">
                                <div class="stat-number">{df.shape[1]}</div>
                                <div class="stat-label">Columns Scanned</div>
                            </div>
                        </div>
                    </div>
                ''', unsafe_allow_html=True)
                
                # Display extracted forms information with modern styling
                st.markdown('''
                    <div class="modern-card">
                        <h3 style="color: #00f5ff; text-align: center; margin-bottom: 2rem;">
                            📋 EXTRACTED BOOKING FORMS
                        </h3>
                    </div>
                ''', unsafe_allow_html=True)
                
                for i, base_data in enumerate(processed_base_data, 1):
                    form_ref = base_data.get('Reference', 'No Reference')
                    form_desc = base_data.get('Description', 'No Description')
                    
                    with st.expander(f"🎯 Form {i}: {form_ref} | {form_desc[:50]}{'...' if len(form_desc) > 50 else ''}"):
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
                
                # Enhanced generate button with modern styling
                st.markdown('''
                    <div class="modern-card">
                        <h3 style="color: #00f5ff; text-align: center; margin-bottom: 1rem;">
                            🚀 FINAL PROCESSING
                        </h3>
                        <p style="color: rgba(255,255,255,0.8); text-align: center; margin-bottom: 2rem;">
                            Generate your final Excel file with all processed booking forms
                        </p>
                    </div>
                ''', unsafe_allow_html=True)
                
                # Create Excel output with enhanced button
                if st.button("🚀 Generate Order Processing Sheet", type="primary", use_container_width=True):
                    # Show processing animation
                    with st.spinner("🔄 Generating Excel file..."):
                        output = BytesIO()
                        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                            edited_order.to_excel(writer, sheet_name='Order Details', index=False)
                            
                            # Enhanced formatting
                            workbook = writer.book
                            header_format = workbook.add_format({
                                'bold': True, 
                                'bg_color': '#4F8BF9', 
                                'font_color': 'white',
                                'border': 1,
                                'align': 'center',
                                'valign': 'vcenter'
                            })
                            cell_format = workbook.add_format({
                                'border': 1,
                                'align': 'left',
                                'valign': 'vcenter'
                            })
                            
                            # Format the worksheet
                            worksheet = writer.sheets['Order Details']
                            
                            # Set column widths and formatting
                            for col_num, value in enumerate(edited_order.columns):
                                worksheet.write(0, col_num, value, header_format)
                                col_width = max(15, len(value) + 5, 
                                              max(len(str(edited_order.iloc[row, col_num])) 
                                                  for row in range(min(len(edited_order), 10))) + 3)
                                worksheet.set_column(col_num, col_num, col_width)
                            
                            # Format data cells
                            for row_num in range(1, len(edited_order) + 1):
                                for col_num in range(len(edited_order.columns)):
                                    worksheet.write(row_num, col_num, 
                                                  edited_order.iloc[row_num-1, col_num], cell_format)
                        
                        output.seek(0)
                    
                    # Enhanced success message with download stats
                    total_forms = len(processed_base_data)
                    total_units = sum(int(form.get('Total Units', '0').replace(',', '') or '0') 
                                    for form in processed_base_data 
                                    if form.get('Total Units', '0').replace(',', '').isdigit())
                    
                    st.markdown(f'''
                        <div class="success-message">
                            <h3>🎉 EXCEL FILE GENERATED SUCCESSFULLY!</h3>
                            <div class="stats-container">
                                <div class="stat-item">
                                    <div class="stat-number">{total_forms}</div>
                                    <div class="stat-label">Forms Included</div>
                                </div>
                                <div class="stat-item">
                                    <div class="stat-number">{len(edited_order)}</div>
                                    <div class="stat-label">Data Rows</div>
                                </div>
                                <div class="stat-item">
                                    <div class="stat-number">{total_units:,}</div>
                                    <div class="stat-label">Total Units</div>
                                </div>
                                <div class="stat-item">
                                    <div class="stat-number">{output.getvalue().__len__()//1024}</div>
                                    <div class="stat-label">KB Size</div>
                                </div>
                            </div>
                        </div>
                    ''', unsafe_allow_html=True)
                    
                    # Generate filename from first form or use generic name
                    first_ref = processed_base_data[0].get('Reference', 'multi_form') if processed_base_data else 'multi_form'
                    filename = f"order_processing_{first_ref.upper()}_forms.xlsx"
                    
                    # Enhanced download button
                    st.download_button(
                        label="📥 Download Multi-Form Order Processing Excel",
                        data=output,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="download_button",
                        use_container_width=True
                    )
            else:
                st.markdown('''
                    <div class="warning-message">
                        <h3>⚠️ NO VALID BOOKING FORMS DETECTED</h3>
                        <p>The AI processor couldn't find any valid booking forms in your Excel file.</p>
                    </div>
                ''', unsafe_allow_html=True)
                
                st.markdown('''
                    <div class="modern-card">
                        <h4 style="color: #00f5ff;">💡 OPTIMIZATION TIPS</h4>
                        <div style="color: rgba(255,255,255,0.8); margin: 1rem 0;">
                            <div style="margin: 0.8rem 0;">
                                <span style="color: #00ff80;">✓</span> Ensure your Excel sheet contains 'Booking Form' headers
                            </div>
                            <div style="margin: 0.8rem 0;">
                                <span style="color: #00ff80;">✓</span> Make sure forms have valid data (not #N/A)
                            </div>
                            <div style="margin: 0.8rem 0;">
                                <span style="color: #00ff80;">✓</span> Check that key fields like Description, Reference are filled
                            </div>
                            <div style="margin: 0.8rem 0;">
                                <span style="color: #00ff80;">✓</span> Verify that delivery dates are properly formatted
                            </div>
                        </div>
                    </div>
                ''', unsafe_allow_html=True)
                
                # Enhanced data preview
                with st.expander("🔍 Raw Data Analysis"):
                    col1, col2 = st.columns(2)
                    with col1:
                        st.write("**File Statistics:**")
                        st.write(f"Total Rows: {df.shape[0]}")
                        st.write(f"Total Columns: {df.shape[1]}")
                        st.write(f"Non-empty Cells: {df.count().sum()}")
                    with col2:
                        st.write("**Data Preview:**")
                        st.dataframe(df.head(10), use_container_width=True)
                
        except Exception as e:
            st.markdown(f'''
                <div class="error-message">
                    <h3>🚨 PROCESSING ERROR DETECTED</h3>
                    <p>Error: <code>{str(e)}</code></p>
                </div>
            ''', unsafe_allow_html=True)
            
            # Enhanced troubleshooting section
            with st.expander("🔧 Advanced Troubleshooting"):
                import traceback
                st.code(traceback.format_exc(), language="python")
            
            # Modern troubleshooting guide
            st.markdown('''
                <div class="modern-card">
                    <h4 style="color: #ff0080;">🛠️ TROUBLESHOOTING GUIDE</h4>
                    <div style="color: rgba(255,255,255,0.8); margin: 1rem 0;">
                        <div style="margin: 1rem 0;">
                            <h5 style="color: #00f5ff;">📦 Dependencies Check</h5>
                            <code style="background: rgba(0,0,0,0.3); padding: 0.5rem; border-radius: 5px;">
                                pip install xlrd>=2.0.1 openpyxl streamlit pandas
                            </code>
                        </div>
                        <div style="margin: 1rem 0;">
                            <h5 style="color: #00f5ff;">📄 File Format Issues</h5>
                            <ul style="margin-left: 1rem;">
                                <li>Try converting .xlsx to .xls or vice versa</li>
                                <li>Ensure file is not password protected</li>
                                <li>Check for corrupted file data</li>
                            </ul>
                        </div>
                        <div style="margin: 1rem 0;">
                            <h5 style="color: #00f5ff;">🎯 Data Structure</h5>
                            <ul style="margin-left: 1rem;">
                                <li>Verify booking form headers are present</li>
                                <li>Check for proper table structure</li>
                                <li>Ensure delivery dates are formatted correctly</li>
                            </ul>
                        </div>
                    </div>
                </div>
            ''', unsafe_allow_html=True)

    # Add footer
    st.markdown("""
        <div style="margin-top: 5rem; padding: 2rem; text-align: center; border-top: 1px solid rgba(0,245,255,0.2);">
            <p style="color: rgba(255,255,255,0.6); font-size: 0.9rem;">
                Powered by <strong style="color: #00f5ff;">Jolanka AI</strong> | 
                Advanced Excel Processing Technology | 
                © 2025 Jolanka Group
            </p>
        </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
