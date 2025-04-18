import pandas as pd
import helperFunctions
from openpyxl import load_workbook

def main():

    ### Extract
    excel_dataframe = pd.read_excel("Customer Order Payment Summary.xlsx", sheet_name=0, engine='openpyxl')
    excel_dataframe.columns.tolist()

    ### Transform
    unpaid_balance = excel_dataframe['Unpaid Amount'].sum()
    fully_paid_count = (excel_dataframe['Unpaid Amount'] <= 0).sum()
    credit_df = excel_dataframe[excel_dataframe['Unpaid Amount'] < 0]
    total_credits = credit_df['Unpaid Amount'].sum()
    total_payments_received = excel_dataframe['Payment Amount'].sum()
    total_orders = excel_dataframe['OrderID'].nunique()
    percent_fully_paid = round((fully_paid_count / total_orders) * 100, 2)

    # Aging Report
    excel_dataframe['DaysToPay'] = (excel_dataframe['PaymentDate'] - excel_dataframe['OrderDate']).dt.days
    aging_bins = [0, 15, 30, 60, 90, 9999]
    aging_labels = ['0-15 Days', '16-30 Days', '31-60 Days', '61-90 Days', '90+ Days']
    excel_dataframe['Aging Category'] = pd.cut(excel_dataframe['DaysToPay'], bins=aging_bins, labels=aging_labels, right=True)
    aging_report = excel_dataframe['Aging Category'].value_counts().sort_index().reset_index()
    aging_report.columns = ['Aging Range', 'Order Count']

    # Payment Methods
    payment_method_dist = excel_dataframe['PaymentMethod'].value_counts().reset_index()
    payment_method_dist.columns = ['Payment Method', 'Count']

    # KPI Table
    summary_dataframe = pd.DataFrame({
        'Metric': [
            'Total Payments Received',
            'Unpaid Balance',
            'Total Credits',
            '% of Orders Fully Paid'
        ],
        'Value': [
            f"${total_payments_received:,.2f}",
            f"${unpaid_balance:,.2f}",
            f"${total_credits:,.2f}",
            f"{percent_fully_paid}%"
        ]
    })

    excel_dataframe = excel_dataframe.drop(['Email', 'City', 'JoinDate', 'OrderItemID', 'ProductID', 'Category', 'Price', 'Order Value'], axis=1)
    excel_dataframe = excel_dataframe.sort_values(by="Aging Category", ascending=True)

    ### Load
    with pd.ExcelWriter('KPI Summary.xlsx', engine='openpyxl') as writer:
        excel_dataframe.to_excel(writer, sheet_name="Payment Aging Summary", index=False)
        summary_dataframe.to_excel(writer, sheet_name="KPI Summary", index=False)
        aging_report.to_excel(writer, sheet_name="KPI Summary", index=False, startrow=len(summary_dataframe) + 3)
        payment_method_dist.to_excel(writer, sheet_name="KPI Summary", index=False, startrow=len(summary_dataframe) + len(aging_report) + 6)

    workbook = load_workbook("KPI Summary.xlsx")
    worksheet = workbook.active
    for worksheet in workbook.worksheets:
        helperFunctions.auto_adjust_columns(worksheet)

    workbook.save("KPI Summary.xlsx")


if __name__ == "__main__":
    print('Generating KPI Report')
    main()
    print('KPI report complete.')
