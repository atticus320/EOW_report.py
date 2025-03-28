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
    Read the raw Excel file, split the 'Date Time' column into 'Date' and 'Time'.
    
    The output DataFrame will include:
      - Date: Extracted from "Date Time", formatted as mm/dd/yyyy.
      - Time: Extracted from "Date Time" (HH:MM format).
      - Event: Original column.
      - Prior: Original column.
      - Survey: Original column.
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
    
    return clean_df

def write_formatted_excel(df, output_file):
    """
    Write the DataFrame to an Excel file with finalized formatting using xlsxwriter.
    
    The formatting includes:
      - An Excel table with a header.
      - Charles Schwab color scheme: dark blue header (#003876) with white text.
      - Arial font (size 10) throughout.
    """
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        # Write data starting at row 1 (leaving row 0 for the custom header)
        df.to_excel(writer, index=False, sheet_name="Report", startrow=1, header=False)
        
        workbook = writer.book
        worksheet = writer.sheets["Report"]
        
        # Define header format using Charles Schwab colors and Arial font.
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
        
        # Write the headers with the custom format.
        for col_num, header in enumerate(df.columns.values):
            worksheet.write(0, col_num, header, header_format)
        
        # Determine the table range (include header row).
        max_row, max_col = df.shape
        # Convert last column index to a letter (assumes <26 columns)
        last_col_letter = chr(65 + max_col - 1)
        table_range = f"A1:{last_col_letter}{max_row + 1}"
        
        # Add an Excel table (with table style for extra polish)
        worksheet.add_table(table_range, {
            'header_row': True,
            'style': 'Table Style Medium 9',
            'columns': [{'header': col} for col in df.columns.values]
        })
        
        # Optionally, adjust column widths.
        worksheet.set_column('A:A', 12)  # Date column
        worksheet.set_column('B:B', 10)  # Time column
        worksheet.set_column('C:C', 30)  # Event column
        worksheet.set_column('D:D', 10)  # Prior column
        worksheet.set_column('E:E', 10)  # Survey column
        
    print(f"Clean report saved to '{output_file}'")

def main():
    # Define the directory where your weekly EOW files are stored.
    directory = r"M:/Trading/8. Miscellaneous/This Week In Markets"
    
    # Find the latest EOW file in the directory.
    input_file = find_latest_eow_file(directory)
    print(f"Using input file: {input_file}")
    
    # Clean and format the raw data.
    cleaned_df = clean_spreadsheet(input_file)
    
    # Define the output file name so it is saved in the same folder as the raw data.
    output_file = os.path.join(directory, "clean_report.xlsx")
    
    # Write the cleaned data to a new Excel file with finalized formatting.
    write_formatted_excel(cleaned_df, output_file)
    
    # Display a preview in the console.
    print("Final Report Preview:")
    print(cleaned_df.head())

if __name__ == '__main__':
    main()


