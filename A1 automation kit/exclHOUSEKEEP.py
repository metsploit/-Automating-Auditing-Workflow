import xlwings as xw
import pyperclip
import ctypes
import sys
import pyautogui
import os
import time

# Step 1: Read all data from the text file
with open('HOUSEKEEPING.txt', 'r') as file:
    text_data = file.read()  # Read entire file content
# Step 1.1: Check if the text file is empty
if not text_data.strip():  # Check if the text data is empty or contains only whitespace
    # Open the Excel workbook and set it to be visible
    app = xw.App(visible=True)  # Make sure Excel is visible
    wb = app.books.open('UPDTGODLEVEL.xlsm')  # Open the Excel file
    sheet = wb.sheets['HOUSEKEEP']  # Select the sheet you want

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
# text_file_path = 'HOUSEKEEPING.txt'
# copy_text_file_content(text_file_path)

# Step 3: Open the Excel workbook and set it to be visible
app = xw.App(visible=True)  # Make sure Excel is visible
wb = app.books.open('UPDTGODLEVEL.xlsm')  # Open the Excel file
sheet = wb.sheets['HOUSEKEEP']  # Select the sheet you want
# Ensure the sheet is active
sheet.activate()  # Activate the sheet
# Step 4: Select cell A1 and paste the clipboard data
sheet.range('A1').select()  # Select cell A1
sheet.api.Paste()  # Perform a normal paste operation




# Step 4.2: Pause and show a message box to inform the user
ctypes.windll.user32.MessageBoxW(0, "Please perform the 'Text to Columns' operation manually, then click OK.", "Manual Step Required", 1)

# Step 4.1:
# Define the range to check
last_row = 200
rows_to_delete = []  # List to store indices of rows to be deleted

# Loop through rows 1 to 500 in column C
for i in range(1, last_row + 1):
    try:
        cell_value = sheet.range(f'C{i}').value  # Get the cell's value
        print(f"Row {i} in column C: {'Empty' if not cell_value else 'Has data'}")

        # If the cell in column C is empty, add the row index to the delete list
        if not cell_value:
            print(f"Row {i} is empty in column C; marking for deletion.")
            rows_to_delete.append(i)

    except Exception as e:
        print(f"Error accessing row {i}: {e}")

# Now delete all marked rows in reverse order to prevent index shifting
for row in reversed(rows_to_delete):
    try:
        print(f"Deleting row {row} as it was marked for deletion.")
        sheet.api.Rows(row).Delete()
    except Exception as e:
        print(f"Error deleting row {row}: {e}")




# Step 5: Run the macro if required
wb.macro('HOUSEKEEP')()  # Uncomment this line and replace 'ACTV' with the macro name if needed


# # Step 7: Fill 'NIL' in column G up to the last row of data
# for i in range(1, last_row + 1):
#     sheet.range(f'G{i}').value = 'NIL'  # Fill 'NIL' in column G

# Step 8: Save and close the workbook
wb.save()
wb.close()

# Optional: Quit the Excel app if you want to close it
app.quit()
