

class Customers:
    customer_name = ""
    customer_code = ""
    customer_type = ""
    payment_term = ""
    credit_limit = ""
    salesperson = ""
    isObsoleted = ""
     # The init method or constructor
    def __init__(self, title):
        self.title = title
    # setter method to set customer name
    def set_customer_name(self, name):
        self.name = name    

    def get_customer_name(self):
        return self.name