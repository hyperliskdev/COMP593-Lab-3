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
        path = sys.argv[1]

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
    # Determine the name and path of the directory to hold the order data files
    # Create the order directory if it does not already exist
    return 

# Split the sales data into individual orders and save to Excel sheets
def process_sales_data(sales_csv, orders_dir):
    # Import the sales data from the CSV file into a DataFrame
    # Insert a new "TOTAL PRICE" column into the DataFrame
    # Remove columns from the DataFrame that are not needed
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