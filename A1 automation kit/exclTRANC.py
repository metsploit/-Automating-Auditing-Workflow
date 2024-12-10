import xlwings as xw
import pyperclip
import ctypes
import sys
import pyautogui
import os
import time

# Step 1: Read all data from the text file
with open('Transcstaff.txt', 'r') as file:
    text_data = file.read()  # Read entire file content
# Step 1.1: Check if the text file is empty
if not text_data.strip():  # Check if the text data is empty or contains only whitespace
    # Open the Excel workbook and set it to be visible
    app = xw.App(visible=True)  # Make sure Excel is visible
    wb = app.books.open('UPDTGODLEVEL.xlsm')  # Open the Excel file
    sheet = wb.sheets['TRANC']  # Select the sheet you want

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
# text_file_path = 'Transcstaff.txt'
# copy_text_file_content(text_file_path)

# Step 3: Open the Excel workbook and set it to be visible
app = xw.App(visible=True)  # Make sure Excel is visible
wb = app.books.open('UPDTGODLEVEL.xlsm')  # Open the Excel file
sheet = wb.sheets['TRANC']  # Select the sheet you want
# Ensure the sheet is active
sheet.activate()  # Activate the sheet
# Step 4: Select cell A1 and paste the clipboard data
sheet.range('A1').select()  # Select cell A1
sheet.api.Paste()  # Perform a normal paste operation

# Step 4.1: Pause and show a message box to inform the user
ctypes.windll.user32.MessageBoxW(0, "Please perform the 'Text to Columns' operation manually, then click OK.", "Manual Step Required", 1)

# # Step 4.2: Wait for user input to continue
# input("Press 'y' and Enter to continue once 'Text to Columns' is done: ")

# Step 5: Run the macro if required
wb.macro('staff')()  # Uncomment this line and replace 'ACTV' with the macro name if needed

last_row = sheet.range('E' + str(sheet.cells.last_cell.row)).end('up').row

# Iterate over each row in column E
for row in range(1, last_row + 1):
    e_value = sheet.range(f'E{row}').value  # Get value in column E
    if e_value == 0:
        # Copy value from column F and paste it into column E
        f_value = sheet.range(f'F{row}').value
        sheet.range(f'E{row}').value = f_value
        print(f"Copied F{row} to E{row}")
    else:
        print(f"Skipped E{row}")

# After copying values, delete the entire column F
sheet.range('F:F').delete()
print("Deleted column F")

# Step 6: Write sequential numbers in column A
# Get the total number of rows with data in column B (assuming data starts from column B)
last_row = sheet.range('B' + str(sheet.cells.last_cell.row)).end('up').row

# Iterate through each row and write sequential numbers in column A
for i in range(1, last_row + 1):
    sheet.range(f'A{i}').value = i  # Write the number in column A

# Step 7: Add DEBIT/CREDIT classification in column F based on values in column E
for i in range(1, last_row + 1):
    value_in_E = sheet.range(f'E{i}').value
    if value_in_E < 0:
        sheet.range(f'F{i}').value = 'DEBIT'
    elif value_in_E > 0:
        sheet.range(f'F{i}').value = 'CREDIT'
    else:
        sheet.range(f'F{i}').value = ''  # Leave empty if zero or not a valid number

# Step 8: Save and close the workbook
wb.save()
wb.close()

# Optional: Quit the Excel app if you want to close it
app.quit()

