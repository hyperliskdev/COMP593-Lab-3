import sys
import os
import datetime
import pandas
import openpyxl
import xlsxwriter
import re

def main():
    sales_csv = get_sales_csv()
    orders_dir = create_orders_dir(sales_csv)
    process_sales_data(sales_csv, orders_dir)

# Get path of sales data CSV file from the command line
def get_sales_csv():
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
        sys.exit(1)

    return

# Create the directory to hold the individual order Excel sheets
def create_orders_dir(sales_csv):
    # Get directory in which sales data CSV file resides
    directory = os.path.dirname(sales_csv)

    # Determine the name and path of the directory to hold the order data files
    todays_date = datetime.date.today().isoformat()
    order_dir_name = f'Orders_{todays_date}'
    
    # Join the order date with the directory the sales csv is in.
    order_dir = os.path.join(directory, order_dir_name)

    # Create the order directory if it does not already exist
    if not os.path.isdir(order_dir):
        os.mkdir(order_dir)
    return order_dir

# Split the sales data into individual orders and save to Excel sheets
def process_sales_data(sales_csv, orders_dir):
    # Import the sales data from the CSV file into a DataFrame
    data = pandas.read_csv(sales_csv)
    # Insert a new "TOTAL PRICE" column into the DataFrame
    data.insert(7, "TOTAL PRICE", data['ITEM QUANTITY'] * data['ITEM PRICE'])
    # Remove columns from the DataFrame that are not needed
    data.drop(columns=["COUNTRY", "POSTAL CODE", "STATE", "ADDRESS", "CITY", "COUNTRY"], inplace=True)
    # Group the rows in the DataFrame by order ID
    for order_id, order_df in data.groupby("ORDER ID"):
    # For each order ID:        
        # Remove the "ORDER ID" column
        order_df.drop(columns=["ORDER ID"], inplace=True)
        # Sort the items by item number
        order_df.sort_values(by="ITEM NUMBER", inplace=True)

        # Append a "GRAND TOTAL" row
        grand_total = order_df['TOTAL PRICE'].sum()
        grand_total_df = pandas.DataFrame({'ITEM PRICE': ['GRAND TOTAL:'], 'TOTAL PRICE': [grand_total]}) 
        order_df = pandas.concat([order_df, grand_total_df])
        # Determine the file name and full path of the Excel sheet
        customer_name = order_df['CUSTOMER NAME'].values[0]
        customer_name = re.sub(r'\W', '', customer_name )
        filename = f'Order{order_id}_{customer_name}.xlsx'
        filepath = os.path.join(orders_dir, filename)
        # Export the data to an Excel sheet
        sheet_name = f'Order {order_id}'

        writer = pandas.ExcelWriter(filepath)

        order_df.to_excel(writer, sheet_name=sheet_name)

        order_workbook = writer.book
        order_worksheet = writer.sheets[sheet_name]

        money_fmt = order_workbook.add_format({'num_format': '$#,##0'})
        order_worksheet.set_column('A:A', 11)
        order_worksheet.set_column('B:B', 11)
        order_worksheet.set_column('C:C', 15)
        order_worksheet.set_column('D:D', 15)
        order_worksheet.set_column('E:E', 15)
        order_worksheet.set_column('F:F', 13)
        order_worksheet.set_column('G:G', 13, money_fmt)
        order_worksheet.set_column('H:H', 10, money_fmt)
        order_worksheet.set_column('I:I', 30)

        writer.close()

        break;
    return



if __name__ == '__main__':  
    main()