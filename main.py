import streamlit as st
import pandas as pd
import tempfile
from datetime import datetime
from openpyxl import load_workbook

def copy_columns(source_df, mapping, additional_values=None):
    dest_df = pd.DataFrame()
    for source_col, dest_col in mapping.items():
        if source_col in source_df.columns:
            dest_df[dest_col] = source_df[source_col]
        else:
            dest_df[dest_col] = additional_values.get(dest_col, "") if additional_values else ""
    return dest_df

def append_close_data(writer, source_df, status, event_type, title_suffix, startrow):
    close_filter = source_df["STATUS"] == status
    close_funds = source_df[close_filter]
    close_df = pd.DataFrame({
        "Fund": close_funds["NAME"],
        "Event Date": close_funds["LATEST INTERIM CLOSE DATE"],
        "Event Type": event_type,
        "Title": close_funds["NAME"] + title_suffix,
        "Close Size": close_funds["LATEST INTERIM CLOSE SIZE (CURR. MN)"]
    })

    if not close_df.empty:
        close_df.to_excel(writer, sheet_name='Events', index=False, header=False, startrow=startrow)
        startrow += len(close_df)

    return startrow

def append_performance_data(writer, source_df, startrow):
    performance_mapping = {
        "NAME": "Fund",
        "": "Performance Date",
        "TARGET IRR - GROSS MIN": "Performance Value (Min)",
        "TARGET IRR - GROSS MAX": "Performance Value (Max)",
        "": "Fund Performance Measurement Type",
        "": "Fund Performance Measurement Unit",
        "": "Performance Source",
        "": "Confidential"
    }
    
    additional_values = {
        "Fund Performance Measurement Unit": "Percentage",
        "Fund Performance Measurement Type": "Target IRR (Gross) (%)"
    }
    
    performance_df = copy_columns(source_df, performance_mapping, additional_values)
    
    performance_df['Performance Source'] = ""
    performance_df['Confidential'] = ""

    # Ensure the correct order of columns and insertion if necessary
    if 'Fund Performance Measurement Type' not in performance_df.columns:
        performance_df.insert(2, 'Fund Performance Measurement Type', "Target IRR (Gross) (%)")
    else:
        performance_df['Fund Performance Measurement Type'] = "Target IRR (Gross) (%)"
    
    if 'Fund Performance Measurement Unit' not in performance_df.columns:
        performance_df.insert(3, 'Fund Performance Measurement Unit', "Percentage")
    else:
        performance_df['Fund Performance Measurement Unit'] = "Percentage"

    if not performance_df.empty:
        performance_df.to_excel(writer, sheet_name='Performances', index=False, header=False, startrow=startrow)
        startrow += len(performance_df)
    
    return startrow

def autofit_columns(writer_path):
    workbook = load_workbook(writer_path)
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        for col in sheet.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2)
            sheet.column_dimensions[column].width = adjusted_width
    workbook.save(writer_path)

def process_file(uploaded_file):
    xls = pd.ExcelFile(uploaded_file)
    sheets = xls.sheet_names
    funds_mapping = {
        "NAME": "Fund",  
        "FUND CURRENCY": "Fund Currency",  
        "VINTAGE / INCEPTION YEAR": "Vintage Year",  
        "STATUS": "Fund Status",  
        "STRATEGY": "Fund Style",  
        "ASSET CLASS": "Asset Class",  
        "FUND STRUCTURE": "Separate Account",  
        "FUND NUMBER (OVERALL)": "Fund Sequence (Total)",  
        "FUND NUMBER (SERIES)": "Fund Series",  
        "LIFESPAN (YEARS)": "Fund Life",  
        "LIFESPAN EXTENSION": "Fund Life Extension",  
        "TARGET SIZE (CURR. MN)": "Target Size (Local Currency m)",  
        "INITIAL TARGET (CURR. MN)": "Initial Target Size (Local Currency m)",  
        "HARD CAP (CURR. MN)": "Hard Cap (Local Currency m)",  
        "OFFER CO-INVESTMENT OPPORTUNITIES TO LPS?": "Fund coinvesting Lps"  
    }

    events_mapping = {
        "NAME": "Fund",  # Source file column A to Destination file column A
        "FUND RAISING LAUNCH DATE": "Event Date",  # Source file column B to Destination file column B
        "": "Event Type",  # Column C doesn't get copied from the source file
        "": "Title",  # Column D doesn't get copied from the source file
        "FINAL CLOSE SIZE (CURR. MN)": "Close Size"  # Source file column AO to Destination file column E
    }

    performances_mapping = {
        "NAME": "Fund",  # Source file column A to Destination file column A
        "": "Performance Date",  # Column B doesn't get copied from the source file
        "": "Fund Performance Measurement Type",  # Column C doesn't get copied from the source file
        "": "Fund Performance Measurement Unit",  # Column D doesn't get copied from the source file
        "TARGET IRR - NET MIN": "Performance Value (Min)",  # Source file column to Destination file column E
        "TARGET IRR - NET MAX": "Performance Value (Max)",  # Source file column to Destination file column F
        "": "Performance Source",  # Column G doesn't get copied from the source file
        "": "Confidential"  # Column H doesn't get copied from the source file
    }

    additional_values_performances = {
        "Fund Performance Measurement Unit": "Percentage",
        "Fund Performance Measurement Type": "Target IRR Net"
    }

    domicile_mapping = {
        "NAME": "Fund",
        "DOMICILE": "Domicile"
    }

    target_geographies_primary_region_mapping = {
        "NAME": "Fund",
        "PRIMARY REGION FOCUS": "Target Geographies Primary Region"
    }

    target_geographies_mapping = {
        "NAME": "Fund",
        "GEOGRAPHIC EXPOSURE": "Fund Target Geography"
    }

    target_sectors_primary_mapping = {
        "NAME": "Fund",
        "INF: PRIMARY SECTOR": "Sector - Primary"
    }

    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
        with pd.ExcelWriter(tmp.name, engine='openpyxl') as writer:
            for sheet in sheets:
                source_df = pd.read_excel(xls, sheet_name=sheet)
                funds_df = copy_columns(source_df, funds_mapping)
                funds_df.to_excel(writer, sheet_name='Funds', index=False)

                events_df = copy_columns(source_df, events_mapping)
                
                # Initialize 'Event Type', 'Title', and 'Close Size' with empty strings
                events_df['Event Type'] = ""
                events_df['Title'] = ""
                events_df['Close Size'] = ""
                
                # Reorder columns for Events tab
                events_df = events_df[['Fund', 'Event Date', 'Event Type', 'Title', 'Close Size']]
                
                # Add the initial Funds data to Events tab with 'Launch' and 'Title'
                events_df['Event Type'] = "Launch"
                events_df['Title'] = events_df['Fund'] + " launches"
                
                events_df.to_excel(writer, sheet_name='Events', index=False)

                # Append additional data to Events tab
                additional_data = {
                    "Fund": source_df["NAME"],
                    "Event Date": source_df["FINAL CLOSE DATE"],
                    "Event Type": "Final Close",
                    "Title": source_df["NAME"] + " reaches final close",
                    "Close Size": source_df["FINAL CLOSE SIZE (CURR. MN)"]
                }
                additional_df = pd.DataFrame(additional_data)

                additional_df.to_excel(writer, sheet_name='Events', index=False, header=False, startrow=len(events_df)+1)

                startrow = len(events_df) + len(additional_df) + 1

                # Append data for various closes
                startrow = append_close_data(writer, source_df, "First Close", "First Close", " reaches first close", startrow)
                startrow = append_close_data(writer, source_df, "Second Close", "Second Close", " reaches second close", startrow)
                startrow = append_close_data(writer, source_df, "Third Close", "Third Close", " reaches third close", startrow)
                startrow = append_close_data(writer, source_df, "Fourth Close", "Fourth Close", " reaches fourth close", startrow)
                startrow = append_close_data(writer, source_df, "Fifth Close", "Fifth Close", " reaches fifth close", startrow)
                startrow = append_close_data(writer, source_df, "Sixth Close", "Sixth Close", " reaches sixth close", startrow)
                startrow = append_close_data(writer, source_df, "Seventh Close", "Seventh Close", " reaches seventh close", startrow)

                # Create Performances tab with initial data
                performances_df = copy_columns(source_df, performances_mapping, additional_values_performances)
                
                # Ensure all necessary columns are present
                for col in ['Fund', 'Performance Date', 'Fund Performance Measurement Type', 'Fund Performance Measurement Unit', 'Performance Value (Min)', 'Performance Value (Max)', 'Performance Source', 'Confidential']:
                    if col not in performances_df.columns:
                        performances_df[col] = ""

                # Initialize non-source columns with appropriate names
                performances_df['Performance Date'] = ""
                performances_df['Fund Performance Measurement Type'] = "Target IRR Net"
                performances_df['Fund Performance Measurement Unit'] = "Percentage"
                performances_df['Performance Source'] = ""
                performances_df['Confidential'] = ""
                
                # Reorder columns for Performances tab
                performances_df = performances_df[['Fund', 'Performance Date', 'Fund Performance Measurement Type', 'Fund Performance Measurement Unit', 'Performance Value (Min)', 'Performance Value (Max)', 'Performance Source', 'Confidential']]
                
                performances_df.to_excel(writer, sheet_name='Performances', index=False)

                # Append additional performance data
                startrow_performance = len(performances_df) + 1
                startrow_performance = append_performance_data(writer, source_df, startrow_performance)

                # Create new tabs with initial data
                domicile_df = copy_columns(source_df, domicile_mapping)
                domicile_df.to_excel(writer, sheet_name='Domicile', index=False)

                target_geographies_primary_region_df = copy_columns(source_df, target_geographies_primary_region_mapping)
                target_geographies_primary_region_df.to_excel(writer, sheet_name='Target_Geographies_Primary_Regi', index=False)

                target_geographies_df = copy_columns(source_df, target_geographies_mapping)
                target_geographies_df.to_excel(writer, sheet_name='Target_Geographies', index=False)

                target_sectors_primary_df = copy_columns(source_df, target_sectors_primary_mapping)
                target_sectors_primary_df.to_excel(writer, sheet_name='Target_Sectors_Primary', index=False)

        # Auto-fit column widths
        autofit_columns(tmp.name)

        # Generate file name with current date and time
        now = datetime.now().strftime("%Y%m%d_%H%M")
        dest_file_name = f"funds_curated_{now}.xlsx"
        dest_file_path = tmp.name

    return dest_file_path, dest_file_name

st.title("Excel Column Copier")
st.write("Upload your source Excel file to create a destination file based on predefined instructions.")

uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

if uploaded_file:
    dest_file_path, dest_file_name = process_file(uploaded_file)
    st.success(f"Destination file '{dest_file_name}' created successfully!")
    st.download_button(
        label="Download Destination File",
        data=open(dest_file_path, "rb"),
        file_name=dest_file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
