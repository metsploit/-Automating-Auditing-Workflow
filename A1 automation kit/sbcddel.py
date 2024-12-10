import xlwings as xw

# Step 1: Open the Excel workbook and set it to be visible
app = xw.App(visible=True)  # Make sure Excel is visible
wb = app.books.open('UPDTGODLEVEL.xlsm')  # Open the Excel file
sheet = wb.sheets['SBCD']  # Select the sheet you want

# Ensure the sheet is active
sheet.activate()  # Activate the sheet
# Step 1.1: Select cell B1 and paste the clipboard data
sheet.range('B1').select()  # Select cell B1
sheet.api.Paste()  # Perform a normal paste operation
# Step 2: Find how many rows have data in column B
last_row = sheet.range('B' + str(sheet.cells.last_cell.row)).end('up').row

# Step 3: Search column B for 'RUN DATE:' and mark rows for deletion
rows_to_delete = []
for row in range(1, last_row + 1):
    cell_value = sheet.range(f"B{row}").value  # Get value in column B
    if cell_value and 'RUN DATE:' in str(cell_value):  # Check if 'RUN DATE:' is in the cell
        rows_to_delete.append(row)

# Step 4: Delete the rows in reverse order to avoid shifting issues
for row in sorted(rows_to_delete, reverse=True):
    sheet.range(f"A{row}").api.EntireRow.Delete()

# Iterate through each row and write sequential numbers in column A
for i in range(1, last_row + 1):
    sheet.range(f'A{i}').value = i  # Write the number in column A
# Step 5: Fill 'NIL' in column G up to the last row of data
for i in range(1, last_row + 1):
    sheet.range(f'E{i}').value = 'NA'  # Fill 'NIL' in column G
# Step 5: Save and close the workbook
wb.save()
wb.close()

# Optional: Quit the Excel app if you want to close it
app.quit()
