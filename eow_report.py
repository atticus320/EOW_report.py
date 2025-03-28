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
        raise FileNotFoundError(f"No files matching the pattern {pattern} were found in {directory}")
    
    latest_file = max(files, key=os.path.getmtime)
    return latest_file

def clean_spreadsheet(input_file):
    """
    Read the raw Excel file, split the 'Date Time' column into 'Date' and 'Time'.
    
    The output DataFrame will include:
    - Date: Extracted from the "Date Time" column.
    - Time: Extracted from the "Date Time" column.
    - Event: Original column.
    - Prior: Original column.
    - Survey: Original column.
    """
    # Load the Excel file
    df = pd.read_excel(input_file)
    print("Original columns:", df.columns.tolist())
    
    # Create a working copy
    df_clean = df.copy()
    
    # Convert "Date Time" to datetime and split into Date and Time columns.
    df_clean['Date Time'] = pd.to_datetime(df_clean['Date Time'], errors='coerce')
    df_clean['Date'] = df_clean['Date Time'].dt.date
    df_clean['Time'] = df_clean['Date Time'].dt.strftime('%H:%M')
    
    # Select only the columns for the final output.
    final_columns = ['Date', 'Time', 'Event', 'Prior', 'Survey']
    clean_df = df_clean[final_columns].copy()
    
    # Sort by Date and Time
    clean_df['Date'] = pd.to_datetime(clean_df['Date'])
    clean_df.sort_values(by=['Date', 'Time'], inplace=True)
    
    return clean_df

def write_formatted_excel(df, output_file):
    """
    Write the DataFrame to an Excel file with finalized formatting using xlsxwriter.
    """
    # Create a Pandas Excel writer using xlsxwriter as the engine.
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name="Report")
        
        workbook = writer.book
        worksheet = writer.sheets["Report"]
        
        # Define a header format.
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'middle',
            'align': 'center',
            'bg_color': '#D3D3D3',  # light gray background
            'border': 1})
        
        # Apply the header format to the header cells.
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
        
        # Optionally adjust the column widths.
        worksheet.set_column('A:A', 12)  # Date column
        worksheet.set_column('B:B', 10)  # Time column
        worksheet.set_column('C:C', 25)  # Event column
        worksheet.set_column('D:D', 10)  # Prior column
        worksheet.set_column('E:E', 10)  # Survey column
        
    print(f"Clean report with formatting saved to '{output_file}'")

def main():
    # Define the directory where your weekly EOW files are stored.
    directory = r"M:/Trading/8. Miscellaneous/This Week In Markets"
    
    # Find the latest EOW file in the directory.
    input_file = find_latest_eow_file(directory)
    print(f"Using input file: {input_file}")
    
    # Clean the data from the input file.
    cleaned_data = clean_spreadsheet(input_file)
    
    # Define the output file name (adjust path if needed).
    output_file = "clean_report.xlsx"
    
    # Write the cleaned data to a new Excel file with finalized formatting.
    write_formatted_excel(cleaned_data, output_file)
    
    # Optionally, print a preview of the cleaned data.
    print("Cleaned Data Preview:")
    print(cleaned_data.head())

if __name__ == '__main__':
    main()

