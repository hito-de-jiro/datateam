import csv
from datetime import datetime

import xlwt
import os


def csv_to_xls(csv_files, xls_file):
    """
    Combines multiple CSV files into a single Excel (.xls) file with separate sheets.

    :param csv_files: List of paths to CSV files.
    :param xls_file: The path to the source Excel file.
    """

    # Create a new Excel workbook
    workbook = xlwt.Workbook()

    for csv_file in csv_files:
        # Check if the file exists
        if not os.path.isfile(csv_file):
            print(f"File {csv_file} not found. Skip it.")
            continue

        # Extract the file name without the extension for the name of the sheet
        sheet_name = os.path.splitext(os.path.basename(csv_file))[0]
        sheet_name = sheet_name[:31]  # Sheets in .xls are limited to 31 characters

        # Create new sheets
        worksheet = workbook.add_sheet(sheet_name)

        with open(csv_file, mode='r', encoding='utf-8') as file:
            reader = csv.reader(file)
            for row_index, row in enumerate(reader):
                for col_index, cell in enumerate(row):
                    worksheet.write(row_index, col_index, cell)

        print(f"Data from {csv_file} is added to the worksheet '{sheet_name[:31]}'.")

    # Save the workbook
    workbook.save(xls_file)
    print(f"Successfully saved Excel file as {xls_file}.")


if __name__ == "__main__":
    # Specify paths to CSV files
    csv_files = [
        'file1.csv',
        'file2.csv',
        'file3.csv'
    ]

    # Specify the path, name and time of creation of the Excel file
    xls_file = f'combined_output_{datetime.now().strftime("%Y-%m-%d_%H-%M-%S")}.xls'

    csv_to_xls(csv_files, xls_file)
