from openpyxl.styles import numbers

# Function to move an excel column after a destination column.
def move_after_column(worksheet, column_to_move, destination_column):
    headers = list(worksheet.columns)
    headers.remove(column_to_move)
    headers.insert((headers.index(destination_column) + 1), column_to_move)
    return worksheet[headers]

# Function to move an excel column before a destination column.
def move_before_column(worksheet, column_to_move, destination_column):
    headers = list(worksheet.columns)
    headers.remove(column_to_move)
    headers.insert((headers.index(destination_column) - 1), column_to_move)
    return worksheet[headers]

# Function to format input columns into currency number format.
def format_currency_columns(worksheet, currency_columns):
    currency_format = '"$"#,##0.00'
    headers = [cell.value for cell in worksheet[1]]
    for index, header in enumerate(headers, start=1):
        if header in currency_columns:
            for cell in worksheet.iter_cols(min_col=index, max_col=index, min_row=2):
                for value in cell:
                    if isinstance(value.value, (int, float)):
                        value.number_format = currency_format

# Function to format input columns into date format.
def format_date_columns(worksheet, date_columns):
    date_format = numbers.FORMAT_DATE_YYYYMMDD2
    headers = [cell.value for cell in worksheet[1]]
    for index, header in enumerate(headers, start=1):
        if header in date_columns:
            for cell in worksheet.iter_cols(min_col=index, max_col=index, min_row=2):
                    for value in cell:
                        if hasattr(value.value, 'strftime'):
                            value.number_format = date_format

# Function to auto adjust the widths of all columns in a worksheet.
def auto_adjust_columns(worksheet):
    for columns in worksheet.columns:
        max_length = 0
        column_letter = columns[0].column_letter
        for cell in columns:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        worksheet.column_dimensions[column_letter].width = max_length + 2 