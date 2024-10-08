import os
import xlwings as xw
from datetime import datetime, timedelta
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
        new_sheet = sheets[-1]
        new_sheet_num = int(last_invoice_sheet.name.split(' ')[-1]) + 1
        new_sheet_name = f"Invoice {new_sheet_num}"
        new_sheet.name = new_sheet_name

        sydney_timezone = pytz.timezone('Australia/Sydney')
        today = datetime.now(sydney_timezone)
        current_year = today.year
        today_formatted = today.strftime("%d/%m/%Y")
        new_sheet.range('C9').number_format = '@'
        new_sheet.range('C9').value = today_formatted
        new_sheet.range('C10').value = new_sheet_num

        num_days, days = input_details(current_year)
        if days is None or num_days is None:
            return False

        print()
        update_excel(new_sheet, days)
        print(f"New sheet created: {new_sheet.name}")

        invoice_folder = os.path.join(os.getcwd(), 'Invoices')
        if not os.path.exists(invoice_folder):
            os.makedirs(invoice_folder)

        pdf_name = f"MRZDesigns - {new_sheet_name}.pdf"
        pdf_path = os.path.join(invoice_folder, pdf_name)

        if os.path.exists(pdf_path):
            new_pdf_name = pdf_name.replace(".pdf", " - Overwritten.pdf")
            os.rename(pdf_path, os.path.join(invoice_folder, new_pdf_name))
            print(f"An Existing PDF was found and was renamed to: {new_pdf_name}")

        new_sheet.api.ExportAsFixedFormat(0, pdf_path)
        print(f"New PDF saved to: {pdf_path}")

        workbook.save()
        return True

    finally:
        app.quit()

def input_details(current_year):
    num_days = input("How many days did you work? (Up to 5): ")
    default_hours_worked = os.getenv('DEFAULT_HOURS_WORKED', '7.5')
    try:
        num_days = int(num_days)
        if not 1 <= num_days <= 5:
            print("Please enter a number between 1 and 5.")
            return None, None
    except ValueError as ve:
        print(ve)
        return None, None

    dates = []
    current_date = None
    for i in range(num_days):
        while True:
            if i == 0 or current_date is None:
                day_month_input = input(f"Enter date for day {i+1} (DD-MM): ")
            else:
                next_date = current_date + timedelta(days=1)
                day_month_input = input(f"Enter date for day {i+1} (DD-MM or press Enter to use {next_date.strftime('%d-%m')}): ")
                if day_month_input == "":
                    current_date = next_date
                    dates.append(current_date.strftime("%d/%m/%Y"))
                    break

            try:
                if '-' in day_month_input:
                    day, month = day_month_input.split('-')
                else:
                    day, month = day_month_input[:2], day_month_input[2:]
                day = int(day)
                month = int(month)
                current_date = datetime(current_year, month, day)
                dates.append(current_date.strftime("%d/%m/%Y"))
                break
            except ValueError as ve:
                print(f"Error: {ve}. Please try again.")

    days = []
    for date in dates:
        hours_input = input(f"Enter hours worked on {date[:5]} (Enter for {default_hours_worked}): ")
        hours_worked = hours_input if hours_input else default_hours_worked
        days.append((date, hours_worked))

    return num_days, days

def update_excel(sheet, days):
    start_row = 16

    ranges_to_clear = ['B16:B20', 'C16:C20', 'D16:D20', 'E16:E20']
    for range_str in ranges_to_clear:
        sheet.range(range_str).clear_contents()

    work_type = os.getenv('WORK_TYPE')
    rate = os.getenv('RATE')

    for index, (date, hours) in enumerate(days):
        row = start_row + index
        sheet.range(f'B{row}').number_format = '@'
        sheet.range(f'B{row}').value = date
        sheet.range(f'C{row}').value = work_type
        sheet.range(f'D{row}').value = hours
        sheet.range(f'E{row}').value = rate