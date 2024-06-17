# Import necessary modules
from win32com import client
import pandas as pd

# Define the file path
file_path = r"C:\Users\OFFSHORE262\Downloads\Invoice Status Report.xls"

# Function to read Excel file and convert to DataFrame
def read_excel_to_df(file_path):
    # Open Microsoft Excel
    excel = client.Dispatch("Excel.Application")
    try:
        # Read Excel File
        workbook = excel.Workbooks.Open(file_path)
        # Select the first worksheet
        worksheet = workbook.Worksheets[1]  # Use 1-based indexing for Worksheets
        # Get the range of used cells in the worksheet
        used_range = worksheet.UsedRange
        # Extract the values from the used range
        data = used_range.Value
        # Convert the data to a DataFrame
        df = pd.DataFrame(list(data))
        # Set the column headers
        df.columns = df.iloc[1]
        df = df.iloc[2:].reset_index(drop=True)
    finally:
        # Close the Excel file
        workbook.Close(SaveChanges=False)
        # Quit the Excel application
        excel.Quit()
    return df

# Call the function and get the DataFrame
df = read_excel_to_df(file_path)
