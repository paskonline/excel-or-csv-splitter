# CSV Splitter (VBA Macro)

A simple, fast VBA tool to automatically split large Excel datasets into manageable, 999-row CSV files. 

If you are dealing with massive data exports (like 50,000+ rows) that crash standard systems or need to be uploaded in smaller batches, this macro handles the heavy lifting directly inside Excel.

## Features
* **Automated Splitting:** Breaks down large sheets into chunks of 999 rows.
* **Header Retention:** Automatically copies the first row (headers) to every single generated CSV file.
* **No External Software Required:** Runs entirely using Excel's built-in VBA engine. No Python or third-party tools needed.
* **Auto-Folder Creation:** Automatically creates a `SplitCSVFiles` folder in the same directory as your workbook to keep your desktop clean.
* **Silent Execution:** Suppresses annoying "Save Changes?" popups to run seamlessly from start to finish.

## How to Use

**Want to test it first?**
I have included sample files in this repository so you can safely test the tool:
* **`SAMPLE_DATASET.xlsm`**: This file already has the macro built-in! Just open it, press `ALT + F11` to view the code, and hit `F5` to run it instantly.
* **`SAMPLE_DATASET.xlsx`**: A raw data file. You can use this to practice the manual setup steps below.

### Step-by-Step Setup:
1. **Download the Code:** Copy the VBA script from `Splitter.ts`.
2. **Open Your Excel File:** Open the large dataset you want to split.
3. **Enable the Developer Tab in Excel**
4. **Open the VBA Editor:** Press `ALT + F11` (Windows) or `Option + F11` (Mac).
5. **Insert a Module:** Click `Insert` > `Module` from the top menu.
6. **Paste and Run:** Paste the code into the blank window and press `F5` (or click the green 'Run' arrow).

*Note: Ensure your data starts on Sheet 1 and that Row 1 contains your headers.*

## Customizing the Row Count
By default, the script splits your file into chunks of 999 rows. If you want a specific row amount, you can easily change this in the code. 

Find **Line 11** in the script and change the number to whatever you need:
```vba
ChunkSize = 999 ' Change 999 to your desired amount
```
## Output
The macro will create a new folder called `SplitCSVFiles` in the exact same location where your original Excel file is saved. Inside, you will find your chunked files named `Part_1.csv`, `Part_2.csv`, etc.

## Important Notes
* You must save your original Excel file as an **Excel Macro-Enabled Workbook (.xlsm)** if you want to keep this code permanently in that specific file.
* The script assumes your data is unbroken. If you have completely blank rows in Column A, it might miscalculate the total row count.
