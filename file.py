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
