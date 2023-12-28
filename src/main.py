import openpyxl
from openpyxl import Workbook
from config import *


# Create a new excel workbook
finance_excel = Workbook()

# Save the workbook at specific path
finance_excel.save(filename=file_path_name)

print(f'Se ha creado el archivo Excel con éxito.')