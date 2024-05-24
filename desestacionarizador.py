import os
from glob import glob

import pandas as pd
import statsmodels.api as sm

def deseasonalize_and_save_as_excel(folder_path):
    """
    Reads all Excel files in the specified folder, deseasonalizes each DataFrame's
    columns, and saves the deseasonalized DataFrame as a new Excel file in a subfolder named "output".

    :param folder_path: str, path to the folder containing Excel files.
    """
    # Get the path to the 'input' subfolder
    input_folder = os.path.join(folder_path, 'input')

    # Create the "output" subfolder if it doesn't exist
    output_folder = os.path.join(folder_path, 'output')
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Get all Excel files in the folder
    excel_files = glob(os.path.join(input_folder, '*.xlsx'))

    # Process each Excel file
    for file in excel_files:
        # Read Excel file
        df = pd.read_excel(file, index_col=0, parse_dates=True)
        
        # Deseasonalize each column using X-13ARIMA-SEATS method
        deseasonalized_dfs = {}
        for column in df.columns:
            serie = df[column]
            res = sm.tsa.x13_arima_analysis(serie, x12path="x13as", outlier=True)
            deseasonalized_dfs[column] = res.seasadj
        
        # Create DataFrame from deseasonalized data
        df_deseasonalized = pd.DataFrame(deseasonalized_dfs)
        
        # Remove hours from the datetime index
        df_deseasonalized.index = df_deseasonalized.index.strftime('%Y-%m-%d')
        
        # Extract file name and extension
        file_name, file_extension = os.path.splitext(os.path.basename(file))
        
        # Construct new file name for the deseasonalized DataFrame
        new_file_name = f"{file_name}_deseasonalized{file_extension}"
        
        # Save deseasonalized DataFrame to new Excel file in the "output" subfolder
        output_path = os.path.join(output_folder, new_file_name)
        df_deseasonalized.to_excel(output_path)
        print(f"New Excel file '{new_file_name}' created with deseasonalized columns in the 'output' subfolder.")

if __name__ == "__main__":
    folder_path = os.path.dirname(os.path.abspath(__file__))  # Use script directory as folder path
    deseasonalize_and_save_as_excel(folder_path)
