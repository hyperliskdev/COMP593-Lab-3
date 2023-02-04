import sys
import os

import pandas
import openpyxl
import xlsxwriter

def main():
    sales_csv = get_sales_csv()
    orders_dir = create_orders_dir(sales_csv)
    process_sales_data(sales_csv, orders_dir)

# Get path of sales data CSV file from the command line
def get_sales_csv():

    path = ""
    # Check whether command line parameter provided
    if len(sys.argv) > 1:
        path = os.path.realpath(sys.argv[1])

       # Check whether provide parameter is valid path of file
        if os.path.exists(path):
            return path
        else:
            print("This file path does not exists")
            exit
    else:
        print("No file path provided")
        exit

    return

# Create the directory to hold the individual order Excel sheets
def create_orders_dir(sales_csv):
    # Get directory in which sales data CSV file resides
    directory = os.path.dirname(sales_csv)

    # Determine the name and path of the directory to hold the order data files
    order_dir = directory + "\Orders"

    # Create the order directory if it does not already exist
    if not os.path.isdir(order_dir):
        os.mkdir(order_dir)
    return order_dir

# Split the sales data into individual orders and save to Excel sheets
def process_sales_data(sales_csv, orders_dir):
    # Import the sales data from the CSV file into a DataFrame
    data = pandas.read_csv(sales_csv)
    # Insert a new "TOTAL PRICE" column into the DataFrame
    data.insert(1, "TOTAL PRICE")
    # Remove columns from the DataFrame that are not needed
    remove_columns = ["ORDER DATE", "ITEM NUMBER", "STATUS"]
    for col in remove_columns:
        data.drop(col)
    # Group the rows in the DataFrame by order ID
    # For each order ID:        
        # Remove the "ORDER ID" column
        # Sort the items by item number
        # Append a "GRAND TOTAL" row
        # Determine the file name and full path of the Excel sheet
        # Export the data to an Excel sheet
        # TODO: Format the Excel sheet
    pass

if __name__ == '__main__':
    main()