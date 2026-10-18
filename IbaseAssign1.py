import openpyxl

file = "bank.xlsx"

# Create Excel file
try:
    wb = openpyxl.load_workbook(file)
    ws = wb.active
except FileNotFoundError:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Name", "Mobile", "Balance"])
    wb.save(file)
    


def add_new():
    wb = openpyxl.load_workbook(file)
    ws = wb.active

    name = input("Enter Your Name: ")
    mobile = input("Enter Mobile Number: ")
    balance = float(input("Enter Opening Balance: "))

    ws.append([name, mobile, balance])
    wb.save(file)
    wb.close()

    print("Account Created Successfully!")


def deposit_money():
    wb = openpyxl.load_workbook(file)
    ws = wb.active

    mobile = input("Enter Mobile Number: ")
    amount = float(input("Enter Amount to Deposit: "))

    for row in range(2, ws.max_row + 1):
        if str(ws.cell(row, 2).value) == mobile:
            ws.cell(row, 3).value += amount
            wb.save(file)
            wb.close()
            print("Money Deposited Successfully!")
            return

    print("Account Not Found!")
    wb.close()


def withdrawal_money():
    wb = openpyxl.load_workbook(file)
    ws = wb.active

    mobile = input("Enter Mobile Number: ")
    amount = float(input("Enter Amount to Withdraw: "))

    for row in range(2, ws.max_row + 1):
        if str(ws.cell(row, 2).value) == mobile:
            balance = ws.cell(row, 3).value

            if amount <= 0:
                print("Enter a Valid Amount!")
            elif amount <= balance:
                ws.cell(row, 3).value = balance - amount
                wb.save(file)
                print("Money Withdrawn Successfully!")
            else:
                print("Insufficient Balance!")

            wb.close()
            return

    print("Account Not Found!")
    wb.close()


def view_balance():
    wb = openpyxl.load_workbook(file)
    ws = wb.active

    mobile = input("Enter Mobile Number: ")

    for row in range(2, ws.max_row + 1):
        if str(ws.cell(row, 2).value) == mobile:
            print("Name:", ws.cell(row, 1).value)
            print("Current Balance: Rs.", ws.cell(row, 3).value)
            wb.close()
            return

    print("Account Not Found!")
    wb.close()


def menu():
    while True:
        print("\nWELCOME TO THE BANK..!")
        print("1. Add New Account")
        print("2. Deposit Money")
        print("3. Withdraw Money")
        print("4. View Balance")
        print("5. Exit")

        choice = input("Enter Your Choice: ")

        if choice == "1":
            add_new()
        elif choice == "2":
            deposit_money()
        elif choice == "3":
            withdrawal_money()
        elif choice == "4":
            view_balance()
        elif choice == "5":
            print("Thank You!")
            break
        else:
            print("Invalid Choice!")


menu()