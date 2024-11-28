import pandas as pd
from openpyxl import load_workbook
from datetime import datetime
import os 
def columninitialize(columns,file_name):
    empty_data = pd.DataFrame(columns=columns)
    empty_data.to_excel(file_name, index=False)

def addingtorow(new_row,file_name):
    try:
        workbook = load_workbook(file_name)
        sheet = workbook.active
        sheet.append(new_row)
        workbook.save(file_name)
    except FileNotFoundError:
        print(f"The file '{file_name}' does not exist. Please create it first.")
def promt():
    print()
    print("Fill The Expenses Details :")
    new_list=[]
    date=input(f"Enter the date :")
    format = "%d-%m-%Y"
    res=True
    try:
        res = bool(datetime.strptime(date, format))
    except ValueError:
        res = False
    if(not(res)):
        print("enter in this formate :" + format)
        promt()
    else:
        categary=input("Enter the catagary :")
        amount=int(input("Enter the amount :"))
        description=input("Enter the description :")
        new_list.append(date)
        new_list.append(categary)
        new_list.append(amount)
        new_list.append(description)
        addingtorow(new_list,file_name)

def view_expenses(file_name):
    data=pd.read_excel(file_name)
    return data

def analys(data):
    if data.empty:
        print("No expenses recorded yet.")
        return
    try:
        print("--------Category Wise-summery---------")
        catagorys=data.groupby('Category')['Amount'].sum()
        print(catagorys)
        print()
        print()
        print("--------Date Wise-summery----------")
        data['Date'] = pd.to_datetime(data['Date'])
        monthly_summary = data.groupby(data['Date'].dt.to_period('M'))['Amount'].sum()
        print(monthly_summary)
        print()
        print()
    except:
        print("------------Please Renter The Data-----------")
columns = ["Date","Category","Amount", "Description",]
file_name = 'expenses.xlsx'
columninitialize(columns,file_name)

while True:
    print("\n--- Expense Tracker Menu ---")
    print("1. Add Expense")
    print("2. View Expenses")
    print("3. Analyze Expenses")
    print("4. Exit")
    option=int(input())
    if(option==1):
        promt()
    elif(option==2):
        print(view_expenses(file_name))
    elif(option==3):
        analys(view_expenses(file_name))
    else:
        break

print("Thanks for your Time")