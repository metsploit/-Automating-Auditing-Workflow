import pyperclip
import tkinter as tk
from tkinter import simpledialog

# Function to process the pasted text
def process_text(text):
    lines = text.splitlines()
    last_run_date = None
    result = []
    
    for line in lines:
        if line.startswith("RUN DATE:"):
            # Extract the date from the RUN DATE line
            last_run_date = line.split()[2]
            result.append(line)  # Keep RUN DATE line unchanged
        else:
            # Replace the line with the last RUN DATE if it's a numbered line
            if line[0].isdigit():
                result.append(last_run_date)
            else:
                result.append(line)  # For any other line (just in case)
    
    return "\n".join(result)

# Function to create a pop-up for user input
def get_user_input():
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    user_input = simpledialog.askstring("Input", "Please paste the text here:")
    return user_input

# Get the user's input from the pop-up
input_text = get_user_input()

# Process the input text
if input_text:
    processed_text = process_text(input_text)
    
    # Copy the processed text to clipboard
    pyperclip.copy(processed_text)
    print("Processed text has been copied to your clipboard.")
else:
    print("No input provided.")
