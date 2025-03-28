import os
import glob
import pandas as pd

def find_latest_eow_file(directory):
    """
    Search the given directory for Excel files with "EOW" in the filename 
    and return the most recent one.
    """
    pattern = os.path.join(directory, "*EOW*.xlsx")
    files = glob.glob(pattern)
    
    if not files:
        raise FileNotFoundError(f"No files matching {pattern} were found in {directory}")
    
    latest_file = max(files, key=os.path.getmtime)
    return latest_file

def clean_spreadsheet(input_file):
    """
    Read the raw Excel file and split the 'Date Time' column into 'Date' and 'Time'.
    The output DataFrame will include:
      - Date: Extracted from "Date Time", formatted as mm/dd/yyyy.
      - Time: Extracted from "Date Time" (HH:MM format).
      - Event: Original column.
      - Prior: Original column.
      - Survey: Original column.
    After sorting, for each day only the first occurrence shows the date and subsequent rows are blank.
    """
    df = pd.read_excel(input_file)
    print("Original columns:", df.columns.tolist())
    
    df_clean = df.copy()
    # Convert "Date Time" to datetime
    df_clean['Date Time'] = pd.to_datetime(df_clean['Date Time'], errors='coerce')
    
    # Create Date and Time columns (Date formatted as mm/dd/yyyy)
    df_clean['Date'] = df_clean['Date Time'].dt.strftime("%m/%d/%Y")
    df_clean['Time'] = df_clean['Date Time'].dt.strftime("%H:%M")
    
    # Select only the columns for the final output.
    final_columns = ['Date', 'Time', 'Event', 'Prior', 'Survey']
    clean_df = df_clean[final_columns].copy()
    
    # Sort by Date and Time
    clean_df = clean_df.sort_values(by=['Date', 'Time'])
    
    # For each date, display the date only on the first occurrence; clear it on subsequent rows.
    clean_df.loc[clean_df.duplicated(subset='Date'), 'Date'] = ""
    
    return clean_df

def write_formatted_excel(df, output_file):
    """
    Write the DataFrame to an Excel file with finalized formatting using xlsxwriter.
    Additional tweaks:
      - A merged title row at the top displaying "Weekly Economic Reports" and the report period.
      - Freeze panes so that title and header remain visible.
      - Charles Schwab header format: dark blue (#003876) with white text, Arial font.
      - Columns B–E (Time, Event, Prior, Survey) are right-aligned.
    """
    # Determine the report period from the non-blank Date values.
    date_values = [cell for cell in df['Date'] if cell != ""]
    if date_values:
        # Convert date strings to datetime objects for comparison.
        dates = pd.to_datetime(date_values, format="%m/%d/%Y")
        report_start = dates.min().strftime("%m/%d/%Y")
        report_end = dates.max().strftime("%m/%d/%Y")
        report_period = f"Period: {report_start} - {report_end}"
    else:
        report_period = ""
    
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        # Write data starting at row 2 (leaving row 0 for the title and row 1 for the header)
        df.to_excel(writer, index=False, sheet_name="Report", startrow=2, header=False)
        
        workbook = writer.book
        worksheet = writer.sheets["Report"]
        
        # Define title format (merged cell) – bold, larger font, dark blue background, white text.
        title_format = workbook.add_format({
            'bold': True,
            'font_name': 'Arial',
            'font_size': 14,
            'align': 'center',
            'valign': 'vcenter',
            'fg_color': '#003876',   # Dark blue
            'font_color': '#FFFFFF'
        })
        # Merge cells A1 to E1 for the title.
        worksheet.merge_range('A1:E1', f"Weekly Economic Reports   {report_period}", title_format)
        
        # Define header format using Charles Schwab colors.
        header_format = workbook.add_format({
            'bold': True,
            'font_name': 'Arial',
            'font_size': 10,
            'align': 'center',
            'valign': 'vcenter',
            'fg_color': '#003876',  # Dark blue
            'font_color': '#FFFFFF',  # White text
            'border': 1
        })
        # Write headers in row 2 (which is Excel row 2+1=3, but our startrow=2 means row index 2)
        for col_num, header in enumerate(df.columns.values):
            worksheet.write(2, col_num, header, header_format)
        
        # Determine table range (including header row, starting from row 3 in Excel terms)
        max_row, max_col = df.shape
        total_rows = max_row + 2  # +2 because we added two extra rows (title and header)
        last_col_letter = chr(65 + max_col - 1)  # Assumes fewer than 26 columns.
        table_range = f"A2:{last_col_letter}{total_rows}"
        
        # Add an Excel table for extra polish.
        worksheet.add_table(table_range, {
            'header_row': True,
            'style': 'Table Style Medium 9',
            'columns': [{'header': col} for col in df.columns.values]
        })
        
        # Adjust column widths.
        worksheet.set_column('A:A', 12)  # Date column
        
        # Define right-aligned format for columns B–E.
        right_align_format = workbook.add_format({'align': 'right', 'font_name': 'Arial', 'font_size': 10})
        worksheet.set_column('B:B', 10, right_align_format)  # Time column
        worksheet.set_column('C:C', 30, right_align_format)  # Event column
        worksheet.set_column('D:D', 10, right_align_format)  # Prior column
        worksheet.set_column('E:E', 10, right_align_format)  # Survey column
        
        # Freeze panes so that rows 0-2 (title and header) remain visible.
        worksheet.freeze_panes(3, 0)
        

    print(f"Clean report saved to '{output_file}'")

def main():
    # Define the directory where your weekly EOW files are stored.
    directory = r"M:/Trading/8. Miscellaneous/This Week In Markets"
    
    # Find the latest EOW file in the directory.
    input_file = find_latest_eow_file(directory)


