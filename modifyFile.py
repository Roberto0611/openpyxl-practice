# Imports
from openpyxl import load_workbook
from openpyxl import Workbook

# Open the workbook
wb = load_workbook(filename = 'excelFiles/sells.xlsx')
sheet = wb["sells"]

# We add a formula
sheet["D2"] = "=B2*C2"

# Save the new file
wb.save("modified.xlsx")