class Customers:
    # The init method or constructor

    def __init__(self, title):
        self.title = title
    # setter method for customer name

    def set_customer_name(self, name):
        self.name = name

    # getter method for customer name
    def get_customer_name(self):
        return self.name

    # setter method for customer code
    def set_customer_code(self, code):
        self.name = code
    # getter method for customer code

    def get_customer_code(self):
        return self._code
    # setter method for customer type

    def set_customer_type(self, type):
        self.name = type
    # getter method for customer type

    def get_customer_type(self):
        return self._type
    # setter method for customer's payment term

    def set_customer_payment_term(self, term):
        self.name = term
    # getter method for customer's payment term

    def get_customer_payment_term(self):
        return self._term
    # setter method for customer's credit limit

    def set_customer_credit_limit(self, limit):
        self.name = limit
    # getter method for customer's credit limit

    def get_customer_credit_limit(self):
        return self._limit
    # setter method for customer's associated salesperson

    def set_customer_salesperson(self, salesperson):
        self.name = salesperson
    # getter method for customer's associated salesperson

    def get_customer_salesperson(self):
        return self._salesperson
    # setter method for customer's obsoleted status

    def set_customer_isObsoleted_status(self, isobsoleted):
        self.name = isobsoleted
    # getter method for customer's obsoleted status

    def get_customer_isObsoleted_status(self):
        return self._isobsoletedname

    def set_customer_aging_balance(self, balance):
        self.name = balance

    # getter method for customer name
    def get_customer_aging_balance(self):
        return self.name        
