import openpyxl
from openpyxl import Workbook
from config import *


# Crear un nuevo libro de trabajo (Workbook)
finance_excel = Workbook()

# Guardar el libro de trabajo en la ruta especificada
finance_excel.save(filename=file_path_name)

print(f'Se ha creado el archivo Excel con éxito.')