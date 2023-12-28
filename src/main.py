import openpyxl
from openpyxl import *
from config import *
import datetime
import os


# INPUT CONTROL
while True:
    input_file_path = input("Input path of your Income and Expense Excel: ")
    
    # Verify the Path and valid input file
    if os.path.exists(input_file_path) and os.path.splitext(input_file_path)[1] in ('.xls', '.xlsx'):
            print("File it's correct")
            break              
    else:
        print("The pathing '"'{}'"' or the file type are incorrect.".format(input_file_path))



# OPEN AND READ THE INPUT EXCEL
expense_income_read = load_workbook(input_file_path)

excel_date = expense_income_read.active[table_start_colum].value

# Verify the date
if isinstance(excel_date, datetime.datetime):
    #Parses the date    
    year = excel_date.year
    month = excel_date.month
    day = excel_date.day

print(excel_date.weekday()) # 0 es Lunes, 6 es domingo --> Para el parseado semanal de las hojas




#CREATE OR COMPLETE THE ANNUAL MONTHLY I&E REPORT
file_name = f"{monthly_income_expense_report} {year}.xlsx"

if os.path.exists(file_name):
    libro_trabajo = load_workbook(file_name)
    hoja = libro_trabajo.active
    # ... 
    
else:
    file_name = Workbook()
    hoja = file_name.active
    # ... 
    

file_name.save(file_name)
