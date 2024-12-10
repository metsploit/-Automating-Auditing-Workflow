import subprocess
import sys
import tkinter as tk
from tkinter import messagebox

# Function to show the confirmation dialog
def ask_user_confirmation():
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    return messagebox.askyesno("Confirmation", 
        "Before proceeding, please open the Excel file UPTDGODLEVEL and manually enable the 'Macro Enabled Content, if you will not enable you face lets of error and dont forget to save the file ans close the excel file after enabaling the content'.\n\n"
        "Did you enable the micro enabled content?")

# Ask for confirmation
proceed = ask_user_confirmation()

# Check user response
if not proceed:
    print("Exiting the script. Please enable macros and run the script again.")
    sys.exit(0)  # Exit the script if the user chooses not to continue

# List your scripts in the desired execution order
scripts = [
    "actvmerge.py",
    # "housekeppmerge.py", i have commented out this codes cause i cant provide txt file to clean these  data 
    # "inopmerge.py",
    # "cifmerge.py",
    "LINKING.py",
    # "excpmerge.py",
    # "ccoddltl.py",
    # "intmerge.py",
    # "sbcdmerge.py",
    # "transmerge.py",
    "exclACTV.py",
    # "exclccod.py",
    # "exclCIF.py",
    # "exclCIFLINK.py",
    # "exclEXCP.py",
    # "exclHOUSEKEEP.py",
    # "exclINOP.py",
    # "exclINTR.py",
    # "exclTRANC.py",
    # "exclSBCD.py",
    # "sbcd.py",
    # "sbcddel.py",
    "aninamerge.py",
    "exclAnina.py",

]

# Iterate through each script and execute it
for script in scripts:
    print(f"Running {script}...")
    result = subprocess.run(['python', script], capture_output=True, text=True)
    print(f"Output:\n{result.stdout}")
    if result.stderr:
        print(f"Error:\n{result.stderr}")
