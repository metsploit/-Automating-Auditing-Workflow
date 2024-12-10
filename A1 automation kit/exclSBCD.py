import xlwings as xw
import pyperclip
import win32com.client as win32
import sys
import ctypes
import pyautogui
import os
import time

# Step 1: Read all data from the text file
with open('SBCD.txt', 'r') as file:
    text_data = file.read()  # Read entire file content
# Step 1.1: Check if the text file is empty
if not text_data.strip():  # Check if the text data is empty or contains only whitespace
    # Open the Excel workbook and set it to be visible
    app = xw.App(visible=True)  # Make sure Excel is visible
    wb = app.books.open('UPDTGODLEVEL.xlsm')  # Open the Excel file
    sheet = wb.sheets['SBCD']  # Select the sheet you want

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
pyperclip.copy(text_data)  # Copy the text data to the clipboard
# def copy_text_file_content(file_path):
#     # Open the text file with the default application (notepad, etc.)
#     os.startfile(file_path)  # This will open the file in the default text editor
#     time.sleep(1)  # Give it a moment to open

#     # Simulate Ctrl + A to select all text
#     pyautogui.hotkey('ctrl', 'a')  # Select all
#     time.sleep(0.5)  # Brief pause before copying
#     pyautogui.hotkey('ctrl', 'c')  # Copy selected text'
#     time.sleep(0.5)  # Brief pause after copying
#     pyautogui.hotkey('alt', 'f4')  # Close the active window

# # Path to your text file
# text_file_path = 'SBCD.txt'
# copy_text_file_content(text_file_path)



# Step 3: Open the Excel workbook and set it to be visible
app = xw.App(visible=True)  # Make sure Excel is visible
wb = app.books.open('UPDTGODLEVEL.xlsm')  # Open the Excel file
sheet = wb.sheets['SBCD']  # Select the sheet you want
# Ensure the sheet is active
sheet.activate()  # Activate the sheet
# Step 4: Select cell A1 and paste the clipboard data
sheet.range('A1').select()  # Select cell A1
sheet.api.Paste()  # Perform a normal paste operation

# Step 5: Run the macro if required
wb.macro('SBCD')()  # Uncomment this line and replace 'ACTV' with the macro name if needed

# Step 6: Copy entire column B of the current sheet
column_b_values = sheet.range('B1:B' + str(sheet.cells.last_cell.row)).value  # Get all values in column B
column_b_text = "\n".join(str(value) for value in column_b_values if value is not None)  # Combine values with newlines
pyperclip.copy(column_b_text)  # Copy to clipboard

wb.save()
wb.close()

# Optional: Quit the Excel app if you want to close it
app.quit()
