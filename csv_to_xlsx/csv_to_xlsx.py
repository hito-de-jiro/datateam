import csv
from datetime import datetime

from openpyxl import Workbook
import os


def csv_to_xlsx(csv_files, xlsx_file):
    """
    Combines multiple CSV files into a single Excel (.xlsx) file with separate sheets.

    :param csv_files: List of paths to CSV-files.
    :param xlsx_file: The path to the source Excel file.
    """

    # Create a new Excel workbook
    workbook = Workbook()

    for index, csv_file in enumerate(csv_files):
        # Check if the file exists
        if not os.path.isfile(csv_file):
            print(f"File {csv_file} not found. Skip it.")
            continue

        # Extract the file name without the extension for the name of the sheet
        sheet_name = os.path.splitext(os.path.basename(csv_file))[0]

        if index == 0:
            # Use the first sheet by default
            worksheet = workbook.active
            worksheet.title = sheet_name[:31]
        else:
            # Create new sheets
            worksheet = workbook.create_sheet(title=sheet_name[:31])

        with open(csv_file, mode='r', encoding='utf-8') as file:
            reader = csv.reader(file)
            for row in reader:
                worksheet.append(row)

        print(f"Data from {csv_file} is added to the worksheet '{sheet_name[:31]}'.")

    # Save the workbook
    workbook.save(xlsx_file)
    print(f"Successfully saved Excel file as {xlsx_file}.")

#


if __name__ == "__main__":
    # Specify paths to CSV files
    csv_files = [
        'file1.csv',
        'file2.csv',
        'file3.csv'
    ]

    # Specify the path, name and time of creation of the Excel file
    xlsx_file = f'combined_output_{datetime.now().strftime("%Y-%m-%d_%H-%M-%S")}.xlsx'

    csv_to_xlsx(csv_files, xlsx_file)
