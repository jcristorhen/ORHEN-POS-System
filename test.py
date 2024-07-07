import datetime
import openpyxl
import os
import webbrowser

# Global variables for Excel workbook and worksheet
EXCEL_FILE = 'buyer_data.xlsx'
EXCEL_SHEET = 'Buyer Orders'
WORKING_CLIENTS_FILE = 'working_clients.xlsx'
INVENTORY_FILE = 'inventory.txt'

# User credentials and roles
USERS = {
    'admin1': {'password': 'l03e1t3', 'role': 'Administrator'}
}

# Function to authenticate user login
def login(role=None):
    while True:
        username = input("Enter username: ").strip()
        password = input("Enter password: ").strip()

        if role == 'Administrator':
            if username in USERS and USERS[username]['password'] == password:
                return username, 'Administrator'
            else:
                print("Invalid username or password. Please try again.")
        elif role == 'Working client':
            try:
                wb = openpyxl.load_workbook(WORKING_CLIENTS_FILE)
                sheet = wb.active

                for row in sheet.iter_rows(values_only=True):
                    if username == row[1] and password == row[2]:
                        return username, 'Working client'

                print("You are not registered. Please contact your admin. Thank you!")
            except FileNotFoundError:
                print("Working clients database not found.")
                return None, None
        else:
            print("Unknown role specified.")
            return None, None

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
        print(f"Added '{item_name}' to cart. Total: PHP {total[0]:,.2f}")
    else:
        print(f"Item with number {item_number} not found in inventory.")

# Function to remove item from cart
def remove_from_cart(item_index, cart, total):
    if not cart:
        print("Cart is empty. Nothing to remove.")
        return
    
    cart_items = list(cart.keys())
    if 1 <= item_index <= len(cart_items):
        item_name = cart_items[item_index - 1]
        if cart[item_name]['quantity'] > 0:
            cart[item_name]['quantity'] -= 1
            total[0] -= cart[item_name]['price']
            if cart[item_name]['quantity'] <= 0:
                del cart[item_name]
            print(f"Removed one '{item_name}' from cart. Total: PHP {total[0]:,.2f}")
        else:
            print(f"No '{item_name}' in the cart to remove.")
    else:
        print("Invalid item number. Please select a valid item number from the cart.")

# Function to clear all items from cart
def clear_cart(cart, total):
    cart.clear()
    total[0] = 0.0
    print("Cart cleared.")

# Function to generate receipt and store buyer data in Excel
def generate_receipt(cart, total, buyer_id):
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print("\n===== Receipt =====")
    print(f"Date/Time: {timestamp}")
    print("-------------------")
    for item, details in cart.items():
        print(f"{item}: {details['quantity']}pcs. PHP {details['price']} = PHP {details['quantity'] * details['price']:.2f}")
    print("-------------------")
    print(f"Total: PHP {total[0]:,.2f}")
    print("===================")

    # Store buyer data in Excel
    try:
        wb = openpyxl.load_workbook(EXCEL_FILE)
    except FileNotFoundError:
        wb = openpyxl.Workbook()

    if EXCEL_SHEET not in wb.sheetnames:
        wb.create_sheet(EXCEL_SHEET)
        sheet = wb[EXCEL_SHEET]
        sheet.append(['Buyer ID', 'Items Ordered', 'Total'])
    else:
        sheet = wb[EXCEL_SHEET]

    row = (buyer_id, ', '.join([f"{item}: {details['quantity']}pcs." for item, details in cart.items()]), f"PHP {total[0]:,.2f}")
    sheet.append(row)

    # Formatting styles
    left_alignment = openpyxl.styles.Alignment(horizontal='left', vertical='center')
    center_alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center')
    buyer_id_fill = openpyxl.styles.PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")  # Light blue
    header_fill = openpyxl.styles.PatternFill(start_color="2F4F4F", end_color="2F4F4F", fill_type="solid")  # Dark gray
    orders_fill = openpyxl.styles.PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")  # Gray
    price_fill = openpyxl.styles.PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")  # Light red

    # Apply formatting
    for cell in sheet[f"A{sheet.max_row}:C{sheet.max_row}"]:
        for c in cell:
            if c.column == 1:
                c.alignment = left_alignment
                c.fill = buyer_id_fill
            elif c.column == 2:
                c.alignment = left_alignment
                c.fill = orders_fill
            elif c.column == 3:
                c.alignment = center_alignment
                c.fill = price_fill
            c.font = openpyxl.styles.Font(bold=True)

    # Apply header formatting
    for col in sheet.iter_cols(min_row=1, max_row=1, min_col=1, max_col=sheet.max_column):
        for cell in col:
            cell.fill = header_fill
            cell.font = openpyxl.styles.Font(bold=True, color="FFFFFF")  # White font color for headers

    wb.save(EXCEL_FILE)
    wb.close()

# Function to load inventory from file and categorize items
def load_inventory(filename):
    inventory = {}

    try:
        with open(filename, 'r') as file:
            category = None
            for line in file:
                line = line.strip()
                if line:
                    if line.isupper():  # Assume it's a category name
                        category = line
                    else:
                        parts = line.split(',')
                        if len(parts) == 3:
                            item_number, name, price_str = parts
                            price = float(price_str.replace('$', '').strip())
                            if category:
                                inventory[item_number] = {'name': name, 'price': price, 'category': category}
                            else:
                                inventory[item_number] = {'name': name, 'price': price}
                        else:
                            print(f"Ignoring malformed line: {line}")
    except FileNotFoundError:
        print(f"Inventory file '{filename}' not found.")

    return inventory

# Function to initialize Excel file with headers if it doesn't exist
def initialize_excel():
    if not os.path.isfile(EXCEL_FILE):
        wb = openpyxl.Workbook()
        sheet = wb.active
        sheet.title = EXCEL_SHEET
        sheet.append(['Buyer ID', 'Items Ordered', 'Total'])
        wb.save(EXCEL_FILE)
        wb.close()

# Function to get the next available Buyer ID
def get_next_buyer_id():
    try:
        wb = openpyxl.load_workbook(EXCEL_FILE)
        sheet = wb[EXCEL_SHEET]

        if sheet.max_row > 1:
            last_buyer_id = sheet.cell(row=sheet.max_row, column=1).value
            next_buyer_id = last_buyer_id + 1
        else:
            next_buyer_id = 1

        wb.close()
    except FileNotFoundError:
        next_buyer_id = 1

    return next_buyer_id

# Function to display the current cart contents
def display_cart(cart, total):
    print("\nCurrent Cart Contents:")
    if cart:
        for index, (item, details) in enumerate(cart.items(), start=1):
            print(f"{index}. {item}: {details['quantity']}pcs. | PHP {details['price']} each")
        print(f"Total: PHP {total[0]:,.2f}")
    else:
        print("Empty cart.")

# Function to handle inventory addition
def add_inventory(item_number, name, price, category, inventory):
    if item_number in inventory:
        print(f"Item number '{item_number}' already exists in inventory. Use update option to modify.")
    else:
        inventory[item_number] = {'name': name, 'price': price, 'category': category}
        save_inventory(INVENTORY_FILE, inventory)
        print(f"Item '{name}' added to inventory.")

# Function to handle inventory removal
def remove_inventory(item_number, inventory):
    if item_number in inventory:
        del inventory[item_number]
        save_inventory(INVENTORY_FILE, inventory)
        print(f"Item '{item_number}' removed from inventory.")
    else:
        print(f"Item number '{item_number}' not found in inventory.")

# Function to save inventory to file
def save_inventory(filename, inventory):
    try:
        with open(filename, 'w') as file:
            for item_number, details in inventory.items():
                if 'category' in details:
                    file.write(f"{details['category']}\n")
                file.write(f"{item_number},{details['name']}, ${details['price']:.2f}\n")
    except IOError:
        print(f"Error saving inventory to {filename}")

# Main function to run the POS system
def main():
    # Initialize Excel file if not exists
    initialize_excel()

    # Load inventory from file
    inventory = load_inventory(INVENTORY_FILE)

    # Initialize cart and total
    cart = {}
    total = [0.0]

    while True:
        print("\n===== Welcome to the POS System =====")
        print("1. Add item to cart")
        print("2. Remove item from cart")
        print("3. Clear cart")
        print("4. Display cart")
        print("5. Checkout")
        print("6. Administrator Menu")
        print("7. Exit")
        choice = input("Enter your choice: ").strip()

        if choice == '1':
            item_number = input("Enter item number to add to cart: ").strip()
            add_to_cart(item_number, inventory, cart, total)
        elif choice == '2':
            if cart:
                display_cart(cart, total)
                item_index = int(input("Enter item number to remove from cart: ").strip())
                remove_from_cart(item_index, cart, total)
            else:
                print("Cart is empty. Nothing to remove.")
        elif choice == '3':
            clear_cart(cart, total)
        elif choice == '4':
            display_cart(cart, total)
        elif choice == '5':
            if cart:
                buyer_id = get_next_buyer_id()
                generate_receipt(cart, total, buyer_id)
                cart.clear()
                total[0] = 0.0
            else:
                print("Cart is empty. Nothing to checkout.")
        elif choice == '6':
            username, role = login('Administrator')
            if role == 'Administrator':
                administrator_menu(inventory)
        elif choice == '7':
            print("Thank you for using the POS system.")
            break
        else:
            print("Invalid choice. Please enter a number from 1 to 7.")

# Function for administrator menu
def administrator_menu(inventory):
    while True:
        print("\n===== Administrator Menu =====")
        print("1. Add item to inventory")
        print("2. Remove item from inventory")
        print("3. View inventory")
        print("4. Open buyer data Excel")
        print("5. Exit to main menu")
        admin_choice = input("Enter your choice: ").strip()

        if admin_choice == '1':
            item_number = input("Enter item number to add to inventory: ").strip()
            name = input("Enter item name: ").strip()
            price = float(input("Enter item price: ").strip())
            category = input("Enter item category (optional): ").strip()
            add_inventory(item_number, name, price, category, inventory)
        elif admin_choice == '2':
            item_number = input("Enter item number to remove from inventory: ").strip()
            remove_inventory(item_number, inventory)
        elif admin_choice == '3':
            print("\nCurrent Inventory:")
            for item_number, details in inventory.items():
                if 'category' in details:
                    print(f"{item_number}: {details['name']} (${details['price']:.2f}) - {details['category']}")
                else:
                    print(f"{item_number}: {details['name']} (${details['price']:.2f})")
        elif admin_choice == '4':
            try:
                webbrowser.open_new(EXCEL_FILE)
            except Exception as e:
                print(f"Failed to open Excel file: {e}")
        elif admin_choice == '5':
            print("Exiting administrator menu.")
            break
        else:
            print("Invalid choice. Please enter a number from 1 to 5.")

if __name__ == "__main__":
    main()
