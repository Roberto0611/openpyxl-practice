# Imports
from openpyxl import load_workbook
from openpyxl import Workbook

# Open the workbook
wb = load_workbook(filename = 'excelFiles/sells.xlsx')
sheet = wb["sells"]

# Initialize the variable
numberOfProducts = 0;

# Iterate over the cells starting on 2 (the first row is the header)
for cell in sheet['A'][1:]: # When using [A] we get all the cells in A, and using 1: we select from the second element to the last one
    if cell.value is not None: 
        numberOfProducts += 1

# Print the number of products
print(f'el numero de productos es: {numberOfProducts}')