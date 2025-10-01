# RFID to Excel Mapper

This project allows you to scan RFID cards, capture their UID, and map them to students or staff in an Excel file.

---

## Setup Instructions

1. Place your Excel file (`.xlsx`) and the Python script (`new.py`) in the **same local directory**.  
2. Open the `new.py` file and update the following lines according to your Excel file:

```python
# Path to your Excel file
file_path = r"C:\Users\<your-username>\Desktop\MAIN READER LIBRARY\final.xlsx"

# Column number in which RFID UIDs are stored (adjust as per your sheet)
rfid_column = 16
```

3. Save the script and run it to start mapping RFID UIDs to your Excel data.

---

## Required Python Packages

Before running the program, install the required libraries using pip:

```bash
pip install keyboard
pip install openpyxl
```

---

## Notes
- Keep your Excel file properly structured so that the `rfid_column` corresponds to the UID column.  
- Run the script with administrator privileges if `keyboard` requires it.  
