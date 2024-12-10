import os
import re
from tkinter import filedialog, Tk

# Define the cleaning function
def clean_text(text):
    # Remove unwanted patterns and characters
    text = re.sub(r'\^L', '', text)        # Remove ^L
    text = re.sub(r'NIL REPORT', '', text) # Remove NIL REPORT
    text = re.sub(r'(?<!\d)(?<!\w)-+(?!\w)(?!\d)', '', text)          # Remove dashes
    
    text = re.sub(r'\n+', '\n', text)      # Remove extra newlines
    return text.strip()

# Function to prompt for file selection only
def select_files():
    # Hide the main tkinter window
    root = Tk()
    root.withdraw()

    # Ask for text files to process
    input_files = filedialog.askopenfilenames(
        title="select BR_Report_on_Accounts_Become_Inoperative_to_Operative.txt",
        filetypes=[("Text files", "*.txt")]
    )

    return input_files

# Get the input files
input_files = select_files()

if not input_files:
    print("No files selected. Exiting...")
    exit(1)

# Define the output file path with the name "inoperative.txt" in the same folder as the script
script_dir = os.path.dirname(os.path.abspath(__file__))  # Get the current script's directory
output_file = os.path.join(script_dir, "inoperative.txt")

# Clear the output file to ensure it starts empty
with open(output_file, 'w') as output:
    output.write('')  # Clear the file

# Process each file one by one
for file_path in input_files:
    filename = os.path.basename(file_path)
    print(f"Processing {filename}...")  # Log the file being processed

    # Read the file and skip the first 10 lines
    with open(file_path, 'r') as file:
        lines = file.readlines()  # Read all lines
        print(f"{filename}: Read {len(lines)} lines.")  # Log number of lines read
        lines_to_write = lines[10:]  # Skip the first 10 lines
        print(f"{filename}: Writing {len(lines_to_write)} lines.")  # Log number of lines being written

    # Write the processed lines to the output file
    with open(output_file, 'a') as output:
        output.writelines(lines_to_write)
        output.write('\n')  # Add a blank line after each file's content

print(f"Done! All content is saved to {output_file}")

# Clean the output file
with open(output_file, 'r') as output:
    output_content = output.read()

cleaned_content = clean_text(output_content)

with open(output_file, 'w') as output:
    output.write(cleaned_content)

# Verify the content of the output file
with open(output_file, 'r') as output:
    output_content = output.read()
    print("Preview of final output content:")
    print(output_content[-1000:])  # Print the last 1000 characters of the output for review
