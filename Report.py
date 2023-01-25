import csv
import json
import pandas as pd
import os
from openpyxl import Workbook
from openpyxl.styles import Font
from datetime import date
from datetime import timedelta

# This function gets the most recent week days, input integer 1...7, 1 = Monday...7 = Sunday
def get_most_recent_weekdays(day):
    today = date.today()
    offset = (today.weekday() - (day-1))
    last_weekday = today - timedelta(days=offset)
    return "{} {} {}".format(last_weekday.strftime('%a'), last_weekday.strftime("%b"),last_weekday.strftime("%d"))

# Return the file name which contain keyword
def get_filename(keyword):
    current_dir = os.listdir() 
    for i in current_dir:
        if keyword.lower() in i.lower():
            return str(i)
    return ""


# Function to convert a CSV to JSON
# Takes the file paths as arguments
def make_json(csvFilePath, jsonFilePath, primaryKey):
    # create a dictionary
    data = {}
    # Open a csv reader called DictReader
    with open(csvFilePath, encoding='utf-8') as csvf:
        csvReader = csv.DictReader(csvf)
        # Convert each row into a dictionary
        # and add it to data
        for rows in csvReader:
            # Assuming a column named 'No' to
            # be the primary key
            key = rows[str(primaryKey)]
            data[key] = rows
    # open a json writer, and use the json.dumps()
    # function to dump data
    with open(jsonFilePath, 'w', encoding='utf-8') as jsonf:
        jsonf.write(json.dumps(data, indent=4))

# Return the attributes of customer object


def load_customer_obj(customer_name, customer_data):
    customer_code = customer_data[customer_name]["\ufeff*Customer Code"]
    customer_type = customer_data[customer_name]["Customer Type"]
    payment_term = customer_data[customer_name]["Payment Terms"]
    credit_limit = customer_data[customer_name]["Credit Limit"]
    salesperson = customer_data[customer_name]["Salesperson"]
    salesperson_prefix = "{}:".format(str(salesperson.split(":", 1)[0]))
    isObsoleted = customer_data[customer_name]["IsObsoleted"]
    return customer_code, customer_type, payment_term, credit_limit, salesperson, salesperson_prefix, isObsoleted

# Construct customer object using the json file


def construct_customer_dictionary(customers, customerJsonPath):
    # call the make_json function
    make_json(customers, customerJsonPath, "*Customer Name")
    with open(customerJsonPath, "r") as content:
        customer_data = json.loads(content.read())
    return customer_data


def generate_aging_report(filename, workbook, sheet, customercsv, customerjson, aged_receivables_summary, line_num, font_style, font_size):
    customer_data = construct_customer_dictionary(customercsv, customerjson)
    df = pd.DataFrame(pd.read_excel(aged_receivables_summary))
    #starting line
    output_file_row_num = line_num
    num_of_col = len(list(df))
    num_of_row = len(df)

    # the first column of the Aged Receivables Summary, which contain all the customer names
    contact = list(df)[0]
    # the current balance owed by customer
    current = list(df)[1]
    # the balance owed by customer less than a week
    leone_week = list(df)[2]
    # second last column of the file, the balance owed by customer more than the setting weeks
    older = list(df)[num_of_col-2]
    # total balance owed, the last column of the file
    total = list(df)[num_of_col-1]
    cell_font = Font(name=font_style, size=font_size, bold=False)
    number_format = "0.00"
    # loop through the contact list, and start generate report
    for i in range(0, num_of_row):
        customer_name = str(df[contact][i])
        # check if the customer is in customer database
        if (customer_name in customer_data):
            output_file_row_num += 1
            # get the set of customer attributes from customer data
            customer_code, customer_type, payment_term, credit_limit, salesperson, salesperson_prefix, isObsoleted = load_customer_obj(
                customer_name, customer_data)
            # input salesperson
            sheet.cell(row=output_file_row_num,
                       column=1).font = cell_font
            sheet.cell(row=output_file_row_num,
                       column=1).number_format = number_format
            sheet.cell(row=output_file_row_num,
                       column=1).value = salesperson_prefix

            # input customer name
            sheet.cell(row=output_file_row_num, column=2).font = cell_font
            sheet.cell(row=output_file_row_num,
                       column=2).number_format = number_format
            sheet.cell(row=output_file_row_num, column=2).value = customer_name

            # input customer code
            sheet.cell(row=output_file_row_num, column=3).font = cell_font
            sheet.cell(row=output_file_row_num,
                       column=3).number_format = number_format
            sheet.cell(row=output_file_row_num, column=3).value = customer_code

            # input payment term
            sheet.cell(row=output_file_row_num, column=4).font = cell_font
            sheet.cell(row=output_file_row_num,
                       column=4).number_format = number_format
            sheet.cell(row=output_file_row_num, column=4).value = payment_term

            # input credit limit
            sheet.cell(row=output_file_row_num, column=5).font = cell_font
            sheet.cell(row=output_file_row_num,
                       column=5).number_format = number_format
            sheet.cell(row=output_file_row_num, column=5).value = credit_limit

            # input current owed, current owed = current + (<1 week)
            current_owed = df[current][i]
            leone_week_owed = df[leone_week][i]
            sheet.cell(row=output_file_row_num, column=6).font = cell_font
            sheet.cell(row=output_file_row_num,
                       column=6).number_format = number_format
            num = int(current_owed)+int(leone_week_owed)
            sheet.cell(row=output_file_row_num, column=6).value = '({0:.2f})'.format(
                abs(num)) if num < 0 else num

            # input the rest of the balance
            k = 7
            for j in range(3, num_of_col-2):
                c = sheet.cell(row=output_file_row_num, column=k)
                c.font = cell_font
                num = int(df[list(df)[j]][i])
                c.value = '({0:.2f})'.format(abs(num)) if num < 0 else num
                k += 1

            # older balance owed
            sheet.cell(row=output_file_row_num, column=k).font = cell_font
            sheet.cell(row=output_file_row_num,
                       column=k).number_format = number_format
            num = int(df[older][i])
            sheet.cell(row=output_file_row_num, column=k).value = '({0:.2f})'.format(
                abs(num)) if num < 0 else num

            # total balance owed
            sheet.cell(row=output_file_row_num, column=k+1).font = cell_font
            sheet.cell(row=output_file_row_num, column=k +
                       1).number_format = number_format
            num = int(df[total][i])
            sheet.cell(row=output_file_row_num, column=k +
                       1).value = '({0:.2f})'.format(abs(num)) if num < 0 else num
        i += 1
    workbook.save(filename)
