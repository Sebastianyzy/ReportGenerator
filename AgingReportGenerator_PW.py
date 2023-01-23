import csv
import json
import Customers



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


def main():
    customer_name = "a1"
    customer_code = "a010"
    customer_type = "food"
    payment_term = "cod"
    credit_limit = "1000"
    salesperson = "03:avin"
    isObsoleted = "false"
    csvFilePath = r'Customers.csv'
    jsonFilePath = r'Customers.json'

    # Call the make_json function
    make_json(csvFilePath, jsonFilePath, "*Customer Name")
    customer = Customers.Customers(customer_name)
    customer.set_customer_name(customer_name)
    customer.set_customer_code(customer_code)
    
    print(customer.get_customer_code())

if __name__ == "__main__":
    main()
