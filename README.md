# Bulk-PowerPoint-Rename

PowerPoint Renaming & Copying Script
This repository contains two Python scripts to rename and organize PowerPoint files while maintaining folder structure and handling errors.

📌 Features
✅ Extracts first 6 words from slides 1 and 2 for renaming
✅ Removes diacritics & special characters
✅ Maintains folder structure while copying files
✅ Handles duplicate names without overwriting
✅ Logs all operations in process_log.txt
✅ Error recovery script moves failed files automatically

📂 File Structure
📁 CantariBise/ # Source folder with PowerPoint files
📁 CantariBun/ # Destination folder with renamed files
├── process_log.txt # Log file with all operations & errors

🚀 How to Use
1️⃣ Run the Main Script
This script scans CantariBise, renames PowerPoint files, and copies them to CantariBun.

Run this command:
python rename_and_copy.py

📝 Log File:

✅ Success: "Renamed and copied file.ppt to New_Name.ppt"
❌ Error: "ERROR reading file.ppt"
2️⃣ Run the Error Processing Script
If files failed to process, this script moves them to the destination folder, maintaining structure.

Run this command:
python process_errors.py

⚙️ Requirements
📌 Python 3.x installed
📌 Required Libraries:

Install dependencies using:
pip install python-pptx pywin32

🛠 Configuration
Modify these folder paths in both scripts if needed:

src_dir = r"C:\Users\maste\Desktop\CantariBise" # Source
dest_dir = r"C:\Users\maste\Desktop\CantariBun" # Destination

📜 License
This project is open-source and free to use. Feel free to contribute! 🚀
