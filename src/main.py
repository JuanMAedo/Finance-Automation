import openpyxl
from openpyxl import Workbook
from config import *
from pathlib import Path




while True:
    input_file_path = input("Input path of your Income and Expense Excel: ")
    input_file_path = Path(input_file_path)
    
    # Verify the Path and valid input file
    if input_file_path.exists() and input_file_path.suffix in ('.xls', '.xlsx'):
            print("File it's correct")
            break              
    else:
        print("The pathing '"'{}'"' or the file type are incorrect.".format(input_file_path))





# Create a new excel workbook
finance_excel = Workbook()

# Save the workbook at specific path
finance_excel.save(filename=file_path_name)

print(f'Se ha creado el archivo Excel con éxito.')
