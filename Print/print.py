import os

main_folder_path = "C:/Users/User/Desktop/Новая папка"

def process_files(folder_path):
    entries = os.listdir(folder_path)

    for entry in entries:
        entry_path = os.path.join(folder_path, entry)
        
        if os.path.isfile(entry_path):
            if entry.endswith(".xlsx") or entry.endswith(".xls"):
                os.startfile(entry_path, "print")

        # Check if the entry is a folder
        elif os.path.isdir(entry_path):
            # Recursively process files in the subfolder
            process_files(entry_path)

# Start processing files in the main folder
process_files(main_folder_path)
