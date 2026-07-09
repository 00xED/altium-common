import os
import pandas as pd

try:
    # Current directory containing CSV files
    csv_directory = os.getcwd()

    # Output Excel file name
    excel_file = 'parts.xlsx'

    # List all CSV files in the current directory
    csv_files = [file for file in os.listdir(csv_directory) if file.endswith('.csv')]

    # Create a Pandas Excel writer using XlsxWriter as the engine.
    # It will be closed automatically.
    with pd.ExcelWriter(excel_file, engine='xlsxwriter') as excel_writer:
        for csv_file in csv_files:
            csv_path = os.path.join(csv_directory, csv_file)
            sheet_name = os.path.splitext(csv_file)[0]
            df = pd.read_csv(csv_path)
            df.to_excel(excel_writer, sheet_name=sheet_name, index=False)

    print(f'Excel file "{excel_file}" created successfully.')

except Exception as e:
    print(f'Error: {e}')

finally:
    input("Press Enter to continue...")