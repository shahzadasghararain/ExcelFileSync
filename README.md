# Excel File Copier

This script copies all Excel files (`.xls` and `.xlsx`) from a source directory and its subdirectories to a specified destination directory. 

## Features

- Recursively searches through the source directory for Excel files.
- Copies found Excel files to the specified destination directory.
- Retains the original file's metadata (such as the last modified time).

## Requirements

- Python 3.x
- `shutil` and `os` modules (included in Python's standard library)

## Installation

1. Ensure you have Python installed. You can download it from [python.org](https://www.python.org/).

2. Clone this repository or download the script file directly.

## Usage

1. Modify the `source_directory` and `destination_directory` variables in the script to reflect your desired source and destination paths. For example:
    ```python
    source_directory = r"c:\path\to\your\source"
    destination_directory = r"c:\path\to\your\destination"
    ```

2. Run the script:
    ```bash
    python copy_excel_files.py
    ```

3. The script will print a message for each file copied and will notify you when the process is complete.

## Example

```python
import os
import shutil

source_directory = r"c:\source"
destination_directory = r"C:\destination"  # change this path to where you want to copy files

# Loop through the files in the source directory and its subdirectories
for foldername, subfolders, filenames in os.walk(source_directory):
    for filename in filenames:
        if filename.endswith('.xls') or filename.endswith('.xlsx'):  # this checks if the file is an Excel file
            source_path = os.path.join(foldername, filename)
            destination_path = os.path.join(destination_directory, filename)

            # Copy the file to the destination
            shutil.copy2(source_path, destination_path)
            print(f"Copied {filename} from {foldername} to {destination_directory}")

print("Done!")
