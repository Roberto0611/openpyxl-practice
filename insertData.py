# Imports
from openpyxl import Workbook

# Create the book where we are going to work and the sheet
book = Workbook()
sheet = book.active

#InsertData
sheet["A1"] = 5
sheet["A2"] = 10
sheet["B1"] = "Hello World!"

# Save the file
book.save("data.xlsx")
