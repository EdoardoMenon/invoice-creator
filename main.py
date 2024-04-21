from create_invoice import create_new_invoice

def show_menu():
    while True:
        print("\nMenu:")
        print("1. Create new invoice")
        print("2. Exit")
        choice = input("Enter your choice (1-2): ")

        if choice == '1':
            success = create_new_invoice()
            if success:
                print("Invoice created successfully.")
                break
            else:
                print("Failed to create invoice. Please try again.")
        elif choice == '2':
            print("Exiting program...")
            break
        else:
            print("Invalid choice. Please enter a number between 1 and 2.")


if __name__ == '__main__':
    show_menu()