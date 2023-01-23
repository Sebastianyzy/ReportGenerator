import csv
import json
import pandas as pd
import Customers
import os


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

    # Open a json writer, and use the json.dumps()
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
    salesperson_prefix = str(salesperson.split(":", 1)[0])
    isObsoleted = customer_data[customer_name]["IsObsoleted"]
    return customer_code,customer_type,payment_term, credit_limit,salesperson, salesperson_prefix, isObsoleted


def main():
    customerCsvFilePath = r'Customers.csv'
    customerJsonPath = r'Customers.json'
    csvFilePath = r"Planway_Poultry_Inc__-_Aged_Receivables_Summary.xlsx"

    # Call the make_json function
    make_json(customerCsvFilePath, customerJsonPath, "*Customer Name")
    # Construct customer object using the json file
    with open(customerJsonPath, "r") as content:
        customer_data = json.loads(content.read())
    
    
    # customer_name = ""
    # customer_code,customer_type,payment_term, credit_limit,salesperson, salesperson_prefix, isObsoleted = load_customer_obj(customer_name , customer_data)

    df = pd.DataFrame(pd.read_excel(csvFilePath))
    num_of_col = len(list(df))
    num_of_row = len(df)
    num_of_balances = num_of_col - 2

    contact = list(df)[0]
    
    for i in range(0, num_of_row):
        customer_name = str(df[contact][i])
        if(customer_name in customer_data):
            print(i)
        
        i+=1

    os.remove(customerJsonPath)
  

    # df = pd.DataFrame(pd.read_excel(csvFilePath))
    # print(df["Aged Receivables Summary"][6])
    # jsonFilePath = r'Customers.json'
    # cust = "A.C.E. WHOLESALE"
    # with open(jsonFilePath, "r") as content:
    #     customer_data = json.loads(content.read())
    #     # key = list(customer_data.keys())
    #     # value = list(customer_data.values())
    #     for k,v in customer_data.items():
    #         if k == df["Aged Receivables Summary"][6]:
    #             print(list(v.keys())[0])
if __name__ == "__main__":
    main()
