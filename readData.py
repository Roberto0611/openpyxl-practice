# Imports
from openpyxl import load_workbook
from openpyxl import Workbook

# Open the workbook
wb = load_workbook(filename='excelFiles/sells.xlsx')

# Access the specific sheet by name
sheet_ranges = wb['sells']

# Print the value of cell A2
print(sheet_ranges['A2'].value)