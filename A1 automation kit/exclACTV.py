import xlwings as xw
import pyperclip
import win32com.client as win32
import sys
import ctypes
import pyautogui
import os
import time

# Step 1: Read all data from the text file
with open('ACTV.txt', 'r') as file:
    text_data = file.read()  # Read entire file content
# Step 1.1: Check if the text file is empty
if not text_data.strip():  # Check if the text data is empty or contains only whitespace
    # Open the Excel workbook and set it to be visible
    app = xw.App(visible=True)  # Make sure Excel is visible
    wb = app.books.open('UPDTGODLEVEL.xlsm')  # Open the Excel file
    sheet = wb.sheets['ACTV']  # Select the sheet you want

    # Write "All Report is Nill" in cell D5 with font size 20
    cell = sheet.range('D5')
    cell.value = "All Report is Nill"  # Set the message
    cell.font.size = 20  # Set the font size to 20

    # Display a message box to the user and exit
    ctypes.windll.user32.MessageBoxW(0, "The text file is empty. Message written to Excel. Exiting the script.", "Empty File", 0)
    wb.save()  # Save the workbook
    wb.close()  # Close the workbook
    app.quit()  # Quit the Excel application
    sys.exit(0)  # Exit the script
# Step 2: Copy the text data to the clipboard
pyperclip.copy(text_data)  # Copy the text data to the clipboarD

# Step 3: Open the Excel workbook and set it to be visible
app = xw.App(visible=True)  # Make sure Excel is visible
wb = app.books.open('UPDTGODLEVEL.xlsm')  # Open the Excel file
sheet = wb.sheets['ACTV']  # Select the sheet you want
# Ensure the sheet is active
sheet.activate()  # Activate the sheet
# Step 4: Select cell A1 and paste the clipboard data
sheet.range('A1').select()  # Select cell A1
sheet.api.Paste()  # Perform a normal paste operation

# Step 5: Run the macro if required
wb.macro('ACTV')()  # Uncomment this line and replace 'ACTV' with the macro name if needed

# Step 6: Write sequential numbers in column A
# Get the total number of rows with data in column B (assuming data starts from column B)
last_row = sheet.range('B' + str(sheet.cells.last_cell.row)).end('up').row

# Iterate through each row and write sequential numbers in column A
for i in range(1, last_row + 1):
    sheet.range(f'A{i}').value = i  # Write the number in column A

# Step 7: Fill 'NIL' in column G up to the last row of data
for i in range(1, last_row + 1):
    sheet.range(f'G{i}').value = 'NIL'  # Fill 'NIL' in column G

# Step 8: Save and close the workbook
wb.save()
wb.close()

# Optional: Quit the Excel app if you want to close it
app.quit()
