#%%
# Import necessary modules
from win32com import client
import pandas as pd
from sharepoint_utils import get_folder_urls
from sharepoint_utils import connect_to_sharepoint
from sharepoint_utils import upload_dataframe_to_sharepoint
#%%
# Define the file path
file_path = r"C:\Users\HP\somaiya.edu\PyDataNinja - Python\C) Production\XLS to XLSX\Input xls file\Invoice Status Report.xls"
#%%
# Function to read Excel file and convert to DataFrame
def read_excel_to_df(file_path):
    # Open Microsoft Excel
    excel = client.Dispatch("Excel.Application")
    try:
        # Read Excel File
        workbook = excel.Workbooks.Open(file_path)
        # Select the first worksheet
        worksheet = workbook.Worksheets[0]  # Use 1-based indexing for Worksheets
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
#%%
# Call the function and get the DataFrame
df = read_excel_to_df(file_path)

# %%
df
# %%
ctx = connect_to_sharepoint('omkar.sutar@somaiya.edu' , 'Sunfl0wer@1234' , 'https://somaiya0.sharepoint.com/sites/PyDataNinja')
# %%
Upload_URL = '/sites/PyDataNinja/Python/C) Production/XLS to XLSX/ouput csv file'
# %%
upload_dataframe_to_sharepoint(ctx , Upload_URL , df , "Invoice summary report.csv")
# %%
