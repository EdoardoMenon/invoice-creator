import os
import xlwings as xw
from datetime import datetime, timedelta
from dotenv import load_dotenv
import pytz

load_dotenv(override=True)


def get_next_invoice_number(sheets):
    invoice_sheets = [s for s in sheets if s.name.startswith('Invoice ')]
    if not invoice_sheets:
        return 1
    return max(int(s.name.split(' ')[-1]) for s in invoice_sheets) + 1


def get_next_customer_invoice_number(workbook, new_sheet, customer_id):
    max_num = 0
    for sheet in workbook.sheets:
        if sheet.name == new_sheet.name or not sheet.name.startswith('Invoice '):
            continue
        if sheet.range('C11').value == customer_id:
            cust_inv_num = sheet.range('C12').value
            if cust_inv_num is not None:
                try:
                    max_num = max(max_num, int(cust_inv_num))
                except (ValueError, TypeError):
                    pass
    return max_num + 1


def create_new_invoice():
    print("\nWhat type of invoice is this for?")
    print("1. Architecture work")
    print("2. Cleaning work")
    work_type_choice = input("Enter your choice (1-2): ").strip()

    if work_type_choice not in ('1', '2'):
        print("Invalid choice. Please enter 1 or 2.")
        return False

    is_architecture = work_type_choice == '1'

    app = xw.App(visible=False)
    try:
        file_path = os.path.abspath(os.getenv('EXCEL_FILE_PATH', ''))
        templates_path = os.path.abspath(os.getenv('TEMPLATES_FILE_PATH', ''))

        if not os.getenv('EXCEL_FILE_PATH'):
            print("No EXCEL_FILE_PATH found in environment variables.")
            return False
        if not os.getenv('TEMPLATES_FILE_PATH'):
            print("No TEMPLATES_FILE_PATH found in environment variables.")
            return False

        workbook = app.books.open(file_path)
        template_wb = app.books.open(templates_path)

        sheets = workbook.sheets
        new_invoice_num = get_next_invoice_number(sheets)

        template_sheet = template_wb.sheets[0 if is_architecture else 1]
        template_sheet.api.Copy(After=workbook.sheets[-1].api)
        template_wb.api.Close(False)

        new_sheet = workbook.sheets[-1]
        new_sheet.name = f"Invoice {new_invoice_num}"

        sydney_timezone = pytz.timezone('Australia/Sydney')
        today = datetime.now(sydney_timezone)
        today_formatted = today.strftime("%d/%m/%Y")

        new_sheet.range('C9').number_format = '@'
        new_sheet.range('C9').value = today_formatted
        new_sheet.range('C10').value = new_invoice_num

        customer_id = new_sheet.range('C11').value
        customer_invoice_num = get_next_customer_invoice_number(workbook, new_sheet, customer_id)
        new_sheet.range('C12').value = customer_invoice_num

        if is_architecture:
            success = fill_architecture_invoice(new_sheet, today.year)
        else:
            success = fill_cleaning_invoice(new_sheet, today.year)

        if not success:
            new_sheet.delete()
            return False

        print(f"New sheet created: {new_sheet.name}")

        invoice_folder = os.path.join(os.getcwd(), 'Invoices', f'Customer ID - {int(customer_id)}')
        if not os.path.exists(invoice_folder):
            os.makedirs(invoice_folder)

        pdf_name = f"IL BUILDING GROUP - Customer Invoice {customer_invoice_num}.pdf"
        pdf_path = os.path.join(invoice_folder, pdf_name)

        if os.path.exists(pdf_path):
            new_pdf_name = pdf_name.replace(".pdf", " - Overwritten.pdf")
            os.rename(pdf_path, os.path.join(invoice_folder, new_pdf_name))
            print(f"An existing PDF was found and renamed to: {new_pdf_name}")

        new_sheet.api.ExportAsFixedFormat(0, pdf_path)
        print(f"New PDF saved to: {pdf_path}")

        workbook.save()
        return True

    finally:
        app.quit()


def fill_architecture_invoice(sheet, current_year):
    description = sheet.range('C16').value
    rate = sheet.range('E16').value

    num_days, days = input_details(current_year)
    if days is None or num_days is None:
        return False

    sheet.range('B16:E20').clear_contents()

    for index, (date, hours) in enumerate(days):
        row = 16 + index
        sheet.range(f'B{row}').number_format = '@'
        sheet.range(f'B{row}').value = date
        sheet.range(f'C{row}').value = description
        sheet.range(f'D{row}').value = hours
        sheet.range(f'E{row}').value = rate

    return True


def fill_cleaning_invoice(sheet, current_year):
    while True:
        date_input = input("Enter date for the cleaning service (DD-MM): ").strip()
        try:
            if '-' in date_input:
                day, month = date_input.split('-')
            else:
                day, month = date_input[:2], date_input[2:]
            service_date = datetime(current_year, int(month), int(day))
            formatted_date = service_date.strftime("%d/%m/%Y")
            break
        except ValueError as ve:
            print(f"Error: {ve}. Please try again.")

    sheet.range('B16').number_format = '@'
    sheet.range('B16').value = formatted_date
    return True


def input_details(current_year):
    default_hours_worked = os.getenv('DEFAULT_HOURS_WORKED', '7.5')

    num_days = input("How many days did you work? (Up to 5): ")
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
                current_date = datetime(current_year, int(month), int(day))
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
