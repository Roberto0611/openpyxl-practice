# Imports
from openpyxl import Workbook

# Create the book where we are going to work and the sheet
book = Workbook()
sheet = book.active

# Save the file
book.save("prueba.xlsx")