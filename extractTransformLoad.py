import pandas as pd
import helperFunctions
from openpyxl import load_workbook

def main():
    ### Extract
    customers = pd.read_excel("Excel Sheets/customers.xlsx", engine='openpyxl')
    orders = pd.read_excel("Excel Sheets/orders.xlsx", engine='openpyxl')
    payments = pd.read_excel("Excel Sheets/payments.xlsx", engine='openpyxl')
    products = pd.read_excel("Excel Sheets/products.xlsx", engine='openpyxl')
    order_items = pd.read_excel("Excel Sheets/order_items.xlsx", engine='openpyxl')

    ### Transform
    # Rename Columns
    orders = orders.rename(columns={'Amount': 'Order Total'})
    payments = payments.rename(columns={'Amount': 'Payment Amount'})

    # Merge
    order_customer = pd.merge(orders, customers, on='CustomerID', how='left')
    order_customer_payment = pd.merge(order_customer, payments, on='OrderID', how='left')
    full_data = pd.merge(order_customer_payment, order_items, on='OrderID', how='left')
    summary_export = pd.merge(full_data, products, on='ProductID', how='left')

    # Drop Duplicates
    summary_export = summary_export.drop_duplicates(subset='CustomerID')

    # Calculate Sums
    summary_export['Order Value'] = summary_export['Quantity'].fillna(0) * summary_export['Price']
    summary_export['Unpaid Amount'] = summary_export['Order Total'].fillna(0) - summary_export['Payment Amount'].fillna(0)

    # Move Column
    summary_export = helperFunctions.move_after_column(summary_export, 'Order Total', 'Order Value')
    summary_export = helperFunctions.move_after_column(summary_export, 'Payment Amount', 'Order Total')

    # Sort
    summary_export = summary_export.sort_values(by='Order Total', ascending=False)

    ### Load
    with pd.ExcelWriter('Customer Order Payment Summary.xlsx', engine='openpyxl') as writer:
        summary_export.to_excel(writer, index=False, sheet_name='Summary Output')

    # Format Output Using Helper Functions
    workbook = load_workbook("Customer Order Payment Summary.xlsx")
    worksheet = workbook.active
    date_columns = ['OrderDate', 'JoinDate', 'PaymentDate']
    currency_columns = ['Price', 'Order Value', 'Order Total', 'Payment Amount', 'Unpaid Amount']
    for worksheet in workbook.worksheets:
        helperFunctions.format_date_columns(worksheet, date_columns)
        helperFunctions.format_currency_columns(worksheet, currency_columns)
        helperFunctions.auto_adjust_columns(worksheet)

    workbook.save("Customer Order Payment Summary.xlsx")


if __name__ == "__main__":
    print("Generating Customer Order Payment Report...")
    main()
    print("Customer Order Payment Report complete.")