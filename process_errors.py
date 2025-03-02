import os
import shutil

# Define paths
LOG_FILE = r"C:\Users\maste\Desktop\CantariBun\process_log.txt"
SRC_DIR = r"C:\Users\maste\Desktop\CantariBise"
DEST_DIR = r"C:\Users\maste\Desktop\CantariBun"

# Function to move failed files
def move_failed_files():
    if not os.path.exists(LOG_FILE):
        print("Log file not found. No errors to process.")
        return

    with open(LOG_FILE, "r", encoding="utf-8") as log:
        lines = log.readlines()

    for line in lines:
        if "ERROR reading" in line or "ERROR copying" in line:
            try:
                file_path = line.split(" ")[-1].strip()
                if os.path.exists(file_path):
                    relative_path = os.path.relpath(file_path, SRC_DIR)
                    dest_file_path = os.path.join(DEST_DIR, relative_path)

                    dest_folder = os.path.dirname(dest_file_path)
                    os.makedirs(dest_folder, exist_ok=True)

                    shutil.move(file_path, dest_file_path)
                    print(f"Moved {file_path} to {dest_file_path}")
                else:
                    print(f"File not found: {file_path}")
            except Exception as e:
                print(f"Error moving file {file_path}: {e}")

# Main function
def main():
    move_failed_files()
    print("Error processing completed.")

if __name__ == "__main__":
    main()
