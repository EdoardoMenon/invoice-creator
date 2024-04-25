from dotenv import load_dotenv
import xlwings as xw
import os

load_dotenv()

def get_totals():
    app = xw.App(visible=False)
    total_sum = 0
    try:
        file_path = os.getenv('EXCEL_FILE_PATH')
        if not file_path:
            print("No EXCEL_FILE_PATH found in environment variables.")
            return False

        workbook = app.books.open(file_path)
        sheets = workbook.sheets

        for sheet in sheets:
            if 'Invoice' in sheet.name:
                total_value = sheet.range('F21').value
                if total_value is not None:
                    total_sum += total_value
                print(f"Total for {sheet.name}: {total_value}")

        print(f"Grand total from all invoices: {total_sum}")
        return True
    finally:
        app.quit()
