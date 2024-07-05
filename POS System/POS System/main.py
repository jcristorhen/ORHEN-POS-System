import datetime
import openpyxl

# Global variable for Excel workbook and worksheet
EXCEL_FILE = 'buyer_data.xlsx'
EXCEL_SHEET = 'Buyer Orders'

# Function to add item to cart
def add_to_cart(item_number, inventory, cart, total):
    if item_number in inventory:
        item_name = inventory[item_number]['name']
        price = inventory[item_number]['price']
        if item_name in cart:
            cart[item_name]['quantity'] += 1
        else:
            cart[item_name] = {'quantity': 1, 'price': price}
        total[0] += price
        print(f"Added '{item_name}' to cart. Total: ${total[0]}")
    else:
        print(f"Item with number {item_number} not found in inventory.")

# Function to generate receipt and store buyer data in Excel
def generate_receipt(cart, total, buyer_id):
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print("\n===== Receipt =====")
    print(f"Date/Time: {timestamp}")
    print("-------------------")
    for item, details in cart.items():
        print(f"{item}: {details['quantity']}pcs. @ ${details['price']} = ${details['quantity'] * details['price']}")
    print("-------------------")
    print(f"Total: ${total[0]}")
    print("===================")

    # Store buyer data in Excel
    wb = openpyxl.load_workbook(EXCEL_FILE)
    sheet = wb[EXCEL_SHEET]

    row = (buyer_id, ', '.join([f"{item}: {details['quantity']}pcs." for item, details in cart.items()]), total[0])
    sheet.append(row)

    wb.save(EXCEL_FILE)
    wb.close()

# Function to load inventory from file
def load_inventory(filename):
    inventory = {}
    try:
        with open(filename, 'r') as file:
            for line in file:
                if line.strip():  # Check if the line is not empty
                    parts = line.strip().split(',')
                    if len(parts) == 3:
                        item_number, name, price = parts
                        inventory[item_number] = {'name': name, 'price': float(price)}
                    else:
                        print(f"Ignoring malformed line: {line.strip()}")
    except FileNotFoundError:
        print(f"Inventory file '{filename}' not found.")
    return inventory

# Function to initialize Excel file with headers if it doesn't exist
def initialize_excel():
    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet.title = EXCEL_SHEET
    sheet.append(['Buyer ID', 'Items Ordered', 'Total'])
    wb.save(EXCEL_FILE)
    wb.close()

# Main function to run the POS system
def main():
    inventory_file = 'inventory.txt'
    initialize_excel()  # Initialize Excel file with headers if it doesn't exist
    inventory = load_inventory(inventory_file)

    cart = {}
    total = [0.0]  # Using a list to pass total by reference

    buyer_id = 1  # Initialize buyer ID

    print("\nWelcome to the Enhanced POS System")

    while True:
        print("\nInventory:")
        for key, item in inventory.items():
            print(f"{key}. {item['name']} - ${item['price']}")

        item_number = input("\nEnter item number to add to cart (Press 'D' when done, 'exit' to exit): ").strip().upper()

        if item_number == 'D':
            proceed = input("Proceed to tally and print receipt? (yes/no): ").strip().lower()
            if proceed == 'yes':
                if len(cart) > 0:
                    generate_receipt(cart, total, buyer_id)
                    cart = {}
                    total[0] = 0.0
                    buyer_id += 1  # Increment buyer ID for the next buyer
                else:
                    print("Cart is empty. Add items before tallying.")
            elif proceed == 'no':
                continue
            else:
                print("Invalid input. Please enter 'yes' or 'no'.")
        
        elif item_number == 'EXIT':
            confirm_exit = input("Do you want to proceed to exit? (yes/no): ").strip().lower()
            if confirm_exit == 'yes':
                break  # Exit the while loop and end the program
            elif confirm_exit == 'no':
                continue
            else:
                print("Invalid input. Please enter 'yes' or 'no'.")

        elif item_number in inventory:
            add_to_cart(item_number, inventory, cart, total)

        else:
            print("Invalid item number. Please select a valid item number from the inventory.")

    print("Exiting the program. Thank you for using the POS system.")

if __name__ == "__main__":
    main()
