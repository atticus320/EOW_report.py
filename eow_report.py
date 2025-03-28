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

def clean_spreadsheet(input_file, output_file):
    """
    Read the raw Excel file, split the 'Date Time' column into 'Date' and 'Time', 
    and save the cleaned data to a new file.
    
    The output will include the following columns:
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
    
    # Select only the columns to be included in the final output.
    final_columns = ['Date', 'Time', 'Event', 'Prior', 'Survey']
    clean_df = df_clean[final_columns].copy()
    
    # Sort the DataFrame by Date and Time
    clean_df['Date'] = pd.to_datetime(clean_df['Date'])
    clean_df.sort_values(by=['Date', 'Time'], inplace=True)
    
    # Save the cleaned report to a new Excel file
    clean_df.to_excel(output_file, index=False)
    print(f"Clean report saved to '{output_file}'")
    
    return clean_df

def main():
    # Define the directory where your weekly EOW files are stored.
    directory = r"M:/Trading/8. Miscellaneous/This Week In Markets"
    
    # Find the latest EOW file in the directory.
    input_file = find_latest_eow_file(directory)
    print(f"Using input file: {input_file}")
    
    # Define the output file name (adjust path if needed).
    output_file = "clean_report.xlsx"
    
    # Process the file and display a preview of the cleaned data.
    cleaned_data = clean_spreadsheet(input_file, output_file)
    print("Cleaned Data Preview:")
    print(cleaned_data.head())

if __name__ == '__main__':
    main()
