import csv
import json
import pandas as pd
import os
from openpyxl import Workbook
from openpyxl.styles import Font
import Report
# Function to convert a CSV to JSON
# Takes the file paths as arguments



def main():
    CUSTOMERS = r'Customers.csv'
    AGED_RECEIVABLES_SUMMARY = r"Planway_Poultry_Inc__-_Aged_Receivables_Summary.xlsx"
    customerJsonPath = r'Customers.json'

    # create a workbook
    wb = Workbook()
    sheet = wb.active
    output_file_row_num = 1

    # Write title
    sheet.cell(row=output_file_row_num, column=1).font = Font(
        name="Arial", size=20, bold=True)
    sheet.cell(row=output_file_row_num,
               column=1).value = "Aged Receivables Summary [INDIVIDUAL]"
    output_file_row_num += 1

    sheet.cell(row=output_file_row_num, column=1).font = Font(
        name="Arial", size=14, bold=False)
    sheet.cell(row=output_file_row_num,
               column=1).value = "PLANWAY POULTRY INC."
    output_file_row_num += 1

    sheet.cell(row=output_file_row_num, column=1).font = Font(
        name="Arial", size=14, bold=False)
    sheet.cell(row=output_file_row_num,
               column=1).value = "Effective as at EOD Wed Jan 11"
    output_file_row_num += 1

    sheet.cell(row=output_file_row_num, column=1).font = Font(
        name="Arial", size=14, bold=False)
    sheet.cell(row=output_file_row_num, column=1).value = "Ageing by due date"
    output_file_row_num += 1

    # start write aging report
    Report.generate_report("PlanwayAging.xlsx", wb, sheet, CUSTOMERS,
                    customerJsonPath, AGED_RECEIVABLES_SUMMARY, output_file_row_num)

    os.remove(customerJsonPath)


if __name__ == "__main__":
    main()
