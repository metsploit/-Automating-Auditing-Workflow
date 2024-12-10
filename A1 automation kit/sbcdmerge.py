import os
import re
from tkinter import filedialog, Tk

# Define the function to process lines containing "REPORT ID:"
def process_line(line):
    # If line contains "REPORT ID:", keep only the part after 60 characters
    if "REPORT ID:" in line:
        return line[105:]  # Keep everything after the first 60 characters
    return line

# Define the cleaning function
def clean_text(text):
    # Remove unwanted patterns and characters
    text = re.sub(r'\^L', '', text)        # Remove ^L
    text = re.sub(r'NIL REPORT', '', text) # Remove NIL REPORT
    text = re.sub(r'(?<!\d)(?<!\w)-+(?!\w)(?!\d)', '', text)         # Remove dashes
    text = re.sub(r'END OF REPORT', '', text) # Remove END OF REPORT
    text = re.sub(r'\n+', '\n', text)      # Remove extra newlines
    return text.strip()

# Define the function to filter lines based on criteria
def filter_lines(lines):
    filtered_lines = []
    for line in lines:
        # Keep the line if it contains "2.800" or "RUN DATE:" or "date"
        if "2.800" in line or "RUN DATE:" in line or "date" in line:
            filtered_lines.append(line)
    return filtered_lines

# Function to prompt for file selection only
def select_files():
    # Hide the main tkinter window
    root = Tk()
    root.withdraw()

    # Ask for text files to process
    input_files = filedialog.askopenfilenames(
        title="select SB_CD_TERM_Deposit_Opened_DAY_SY7732.txt",
        filetypes=[("Text files", "*.txt")]
    )

    return input_files

# Get the input files
input_files = select_files()

if not input_files:
    print("No files selected. Exiting...")
    exit(1)

# Define the output file path with the name "SBCD.txt" in the same folder as the script
script_dir = os.path.dirname(os.path.abspath(__file__))  # Get the current script's directory
output_file = os.path.join(script_dir, "SBCD.txt")

# Clear the output file to ensure it starts empty
with open(output_file, 'w') as output:
    output.write('')  # Clear the file

# Process each file one by one
for file_path in input_files:
    filename = os.path.basename(file_path)
    print(f"Processing {filename}...")  # Log the file being processed

    # Read the file
    with open(file_path, 'r') as file:
        lines = file.readlines()  # Read all lines
        print(f"{filename}: Read {len(lines)} lines.")  # Log number of lines read

        # Copy only the first line
        lines_to_write = []
        if len(lines) > 0:
            first_line = lines[0]
            lines_to_write.append(first_line)  # Start with the first line

            # Skip the next 10 lines and add the rest
            if len(lines) > 10:
                lines_to_write.extend(lines[11:])

        # Log number of lines being written
        print(f"{filename}: Writing {len(lines_to_write)} lines.")

    # Process each line to handle "REPORT ID:" cases
    lines_to_write = [process_line(line) for line in lines_to_write]

    # Write the processed lines to the output file
    with open(output_file, 'a') as output:
        output.writelines(lines_to_write)
        output.write('\n')  # Add a blank line after each file's content

print(f"Done! All content is saved to {output_file}")

# Read the output file to apply the final filter and cleaning
with open(output_file, 'r') as output:
    output_content = output.read()

# Filter lines based on the presence of "2.800" or specific keywords
filtered_content = '\n'.join(filter_lines(output_content.splitlines()))

# Clean the filtered content
cleaned_content = clean_text(filtered_content)

# Write the cleaned content back to the output file
with open(output_file, 'w') as output:
    output.write(cleaned_content)

# Verify the content of the output file
with open(output_file, 'r') as output:
    output_content = output.read()
    print("Preview of final output content:")
    print(output_content[-1000:])  # Print the last 1000 characters of the output for review
