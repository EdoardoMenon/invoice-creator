from create_invoice import create_new_invoice
from get_invoice_totals import get_totals

def show_menu():
    while True:
        print("\nMenu:")
        print("1. Create new invoice")
        print("2. Get total from all invoices")
        print("3. Exit")
        choice = input("Enter your choice (1-3): ")

        if choice == '1':
            success = create_new_invoice()
            if success:
                print("Invoice created successfully.")
            else:
                print("Failed to create invoice. Please try again.")
        elif choice == '2':
            success = get_totals()
            if success:
                print("Totals retrieved.")
            else:
                print("Failed to get invoice totals. Please try again.")
        elif choice == '3':
            print("Exiting program...")
            break
        else:
            print("Invalid choice. Please enter a number between 1 and 3.")

if __name__ == '__main__':
    show_menu()