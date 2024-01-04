from ctypes import alignment
import datetime
import os
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, PatternFill, Font
from parser_datetime import *
from dotenv import load_dotenv

#SECRETS LOADINGS AT VARIABLES
load_dotenv()
input_file_path = os.getenv('INCOME_EXPENSE_RECORD')
monthly_income_expense_report = os.getenv('MONTHLY_REPORT_NAME')
table_start_colum = os.getenv('START_COLUMN')
table_finish_column = os.getenv('FINISH_COLUMN')

# VERIFY INPUT FILE
if os.path.exists(input_file_path) and os.path.splitext(input_file_path)[1] in ('.xls', '.xlsx'):
        print("File it's correct")         
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

#print(excel_date.weekday()) # 0 es Lunes, 6 es domingo --> Para el parseado semanal de las hojas
#print(excel_date.month)



#CREATE OR COMPLETE THE ANNUAL MONTHLY I&E REPORT
file_name = f"{monthly_income_expense_report} {year}.xlsx"
print(file_name)
if os.path.exists(file_name):
    inc_exp_excel = load_workbook(file_name)
    hoja = inc_exp_excel.active

    
    
    # ... 
    
else:
    inc_exp_excel = Workbook()
    inc_exp_excel.create_sheet("DASHBOARD")
    inc_exp_excel.create_sheet("CATEGORY")
    # CREATE MONTHLY SHEETS
    for month_number, month_name in month_to_name.items():
        month_sheet = inc_exp_excel.create_sheet(month_name)
        inc_exp_excel.active = month_sheet
        
        month_sheet['B2'] = f"EXPENSE"
        cell_B2 = month_sheet['B2']
        cell_B2.alignment = Alignment(horizontal='center')
        cell_B2.fill = PatternFill(start_color="C06FCA", end_color="C06FCA", fill_type="solid")  # Color morado claro
        cell_B2.font = Font(bold=True, name='Calibri') 
        for i in range(5):
            cell = month_sheet.cell(row=3 + i, column=2)  # B3, B4...
            week_text = f"Week {i + 1}"
            cell.value = week_text

            cell.alignment = Alignment(horizontal='center')
            cell.fill = PatternFill(start_color="D78740", end_color="D78740", fill_type="solid")  # Color naranja apagado
            cell.font = Font(bold=True, name='Calibri') 
                
    inc_exp_excel.remove(inc_exp_excel["Sheet"])
    # ... 
    
inc_exp_excel.save(file_name)
