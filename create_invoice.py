import os
import xlwings as xw
from datetime import datetime
from dotenv import load_dotenv
import pytz

load_dotenv()

def create_new_invoice():
    app = xw.App(visible=False)
    try:
        file_path = os.getenv('EXCEL_FILE_PATH')
        if not file_path:
            print("No EXCEL_FILE_PATH found in environment variables.")
            return False

        workbook = app.books.open(file_path)
        sheets = workbook.sheets

        invoice_sheets = [sheet for sheet in sheets if sheet.name.startswith('Invoice ')]
        if not invoice_sheets:
            print("No invoice sheets found following the 'Invoice X' naming convention.")
            return False

        last_invoice_sheet = sorted(invoice_sheets, key=lambda s: int(s.name.split(' ')[-1]))[-1]
        last_invoice_sheet.api.Copy(After=sheets[-1].api)
        new_sheet = sheets[-1]  # New sheet is now the last sheet
        new_sheet_num = int(last_invoice_sheet.name.split(' ')[-1]) + 1
        new_sheet_name = f"Invoice {new_sheet_num}"
        new_sheet.name = new_sheet_name

        sydney_timezone = pytz.timezone('Australia/Sydney')
        today = datetime.now(sydney_timezone)
        current_year = today.year
        today_formatted = today.strftime("%d-%m-%Y")
        new_sheet.range('C9').value = today_formatted
        new_sheet.range('C10').value = new_sheet_num

        num_days, days = input_details(current_year)
        if days is None or num_days is None:
            return False

        update_excel(new_sheet, days)
        print(f"New sheet created: {new_sheet.name}")
        pdf_path = os.path.join(os.getcwd(), f"MRZDesigns - {new_sheet_name}.pdf")
        new_sheet.api.ExportAsFixedFormat(0, pdf_path)
        print("\nSubject: MRZDesigns - Invoice " + str(new_sheet_num) + " - " + str(num_days) + " days")

        workbook.save()
        return True

    finally:
        app.quit()

def input_details(current_year):
    num_days = input("How many days did you work? (Up to 5): ")
    try:
        num_days = int(num_days)
        if not 1 <= num_days <= 5:
            print("Please enter a number between 1 and 5.")
            return None, None
    except ValueError as ve:
        print(ve)
        return None, None

    days = []
    for i in range(num_days):
        day_month_input = input(f"Enter date for day {i+1} (DD-MM): ")
        hours_worked = input(f"Enter hours worked on day {i+1}: ")
        try:
            if '-' in day_month_input:
                day, month = day_month_input.split('-')
            else:
                day, month = day_month_input[:2], day_month_input[2:]
            day = f"{int(day):02d}" 
            month = f"{int(month):02d}"
            date_str = f"{day}-{month}-{current_year}"
            datetime.strptime(date_str, "%d-%m-%Y")
            days.append((date_str, hours_worked))
        except ValueError as ve:
            print(f"Error: {ve}. Please try again.")
            return None, None

    return num_days, days

def update_excel(sheet, days):
    start_row = 16

    # Clear previous data in the specified ranges before entering new data
    ranges_to_clear = ['B16:B20', 'C16:C20', 'D16:D20', 'E16:E20']
    for range_str in ranges_to_clear:
        sheet.range(range_str).clear_contents()

    work_type = os.getenv('WORK_TYPE')
    rate = os.getenv('RATE')

    # Populate the new data into the sheet
    for index, (date, hours) in enumerate(days):
        row = start_row + index
        sheet.range(f'B{row}').value = date 
        sheet.range(f'C{row}').value = work_type
        sheet.range(f'D{row}').value = hours 
        sheet.range(f'E{row}').value = rate
