import os
import re
from tkinter import filedialog, Tk

# Define the cleaning function
def clean_text(text):
    # Remove unwanted patterns and characters
    text = re.sub(r'\^L', '', text)        # Remove ^L
    text = re.sub(r'NIL REPORT', '', text) # Remove NIL REPORT
    text = re.sub(r'(?<!\d)(?<!\w)-+(?!\w)(?!\d)', '', text)  # Remove dashes
    text = re.sub(r'END OF REPORT', '', text)    # Remove END OF REPORT
    text = re.sub(r'\n+', '\n', text)      # Remove extra newlines
    return text.strip()

# Function to remove lines containing "CENTRAL BANK OF INDIA" and the next 9 lines
def remove_central_bank_lines(content):
    lines = content.splitlines()
    filtered_lines = []
    skip_count = 0  # Track the number of lines to skip after "CENTRAL BANK OF INDIA" is found
    
    for line in lines:
        if skip_count > 0:
            skip_count -= 1
            continue
        if "CENTRAL BANK OF INDIA" in line:
            skip_count = 5  # Skip the next 9 lines as well
        else:
            filtered_lines.append(line)
    
    return '\n'.join(filtered_lines)

# Function to prompt for file selection only
def select_files():
    # Hide the main tkinter window
    root = Tk()
    root.withdraw()

    # Ask for text files to process
    input_files = filedialog.askopenfilenames(
        title="select TWO REPORT FROM THE CCOD AND DLTL ",
        filetypes=[("Text files", "*.txt")]
    )

    return input_files

# Get the input files
input_files = select_files()

if not input_files:
    print("No files selected. Exiting...")
    exit(1)

# Define the output file path with the name "CIFlinking.txt" in the same folder as the script
script_dir = os.path.dirname(os.path.abspath(__file__))  # Get the current script's directory
output_file = os.path.join(script_dir, "CCODDLTL.txt")

# Clear the output file to ensure it starts empty
with open(output_file, 'w') as output:
    output.write('')  # Clear the file

# Process each file one by one
for file_path in input_files:
    filename = os.path.basename(file_path)
    print(f"Processing {filename}...")  # Log the file being processed

    # Read the file and skip the first 11 lines
    with open(file_path, 'r') as file:
        lines = file.readlines()  # Read all lines
        print(f"{filename}: Read {len(lines)} lines.")  # Log number of lines read
        # Write all lines after skipping the first 11
        lines_to_write = lines[9:]  # Skip the first 11 lines
        print(f"{filename}: Writing {len(lines_to_write)} lines after initial skip.")  # Log number of lines being written
    # Trim leading and trailing whitespaces from each line to maintain consistency

    with open(output_file, 'a') as output:
        output.writelines(lines_to_write)
        output.write('\n')  # Add a blank line after each file's content

print(f"Done! All content is saved to {output_file}")

# Clean the output file
with open(output_file, 'r') as output:
    output_content = output.read()

cleaned_content = clean_text(output_content)

# Overwrite the output file with the cleaned content
with open(output_file, 'w') as output:
    output.write(cleaned_content)

# Now filter out lines containing "CENTRAL BANK OF INDIA" and the next 9 lines
with open(output_file, 'r') as output:
    final_content = output.read()

filtered_content = remove_central_bank_lines(final_content)

# Overwrite the output file with the filtered content
with open(output_file, 'w') as output:
    output.write(filtered_content)

# Add 8 spaces at the start of the first line
with open(output_file, 'r') as output:
    final_content = output.readlines()

final_content[0] = ' ' * 8 + final_content[0]  # Add 8 spaces at the beginning of the first line

# Write the modified content back to the file
with open(output_file, 'w') as output:
    output.writelines(final_content)

# Verify the content of the output file
with open(output_file, 'r') as output:
    output_content = output.read()
    print("Preview of final output content:")
    print(output_content[-1000:])  # Print the last 1000 characters of the output for review
