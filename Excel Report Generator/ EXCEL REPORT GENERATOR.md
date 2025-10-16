================================================================
                    EXCEL REPORT GENERATOR
================================================================


WHAT IT DOES (In 10 Seconds)
────────────────────────────────────────────────────────────────

Converts Word files (.docx) → Excel files (.xlsx) automatically.

Input:  Word documents with IT events (date, priority, client)
Output: Perfect Excel spreadsheets with formatted data
Time:   1 second per file
Cost:   Free
Privacy: 100% offline (no internet needed)


QUICK START - Complete Setup Instructions
────────────────────────────────────────────────────────────────

PREREQUISITES:
Before running the script, you need:
✓ Python 3.7 or higher
✓ pip (package manager - comes with Python)
✓ Word files (.docx) with event data


STEP 1: Install Python
──────────────────────

macOS (using Homebrew):
brew install python3

Linux (Ubuntu/Debian):
sudo apt-get update
sudo apt-get install python3 python3-pip

Windows:
Download installer from: https://www.python.org/downloads/
Run the installer and follow the prompts
Important: Check the box "Add Python to PATH"


STEP 2: Verify Installation
────────────────────────────

Check Python version:
python3 --version

or on Windows:
python --version

Check pip version:
pip --version

(Both should show version numbers if installed correctly)


STEP 3: Prepare Your Environment
─────────────────────────────────

Create a directory for the script:
mkdir excel-converter
cd excel-converter

Copy files into this directory:
- excel_generator.py (the script)
- requirements.txt (dependencies list)

Place your Word files (.docx) in this same directory


STEP 4: Install Dependencies (Execute ONCE)
────────────────────────────────────────────

macOS & Linux:
pip install -r requirements.txt

or

pip3 install -r requirements.txt

Windows (PowerShell or Command Prompt):
pip install -r requirements.txt

(This installs: python-docx and openpyxl libraries)


STEP 5: Run the Script
──────────────────────

macOS & Linux:
python3 excel_generator.py

or

python excel_generator.py

Windows (PowerShell or Command Prompt):
python excel_generator.py

Wait for message: ✅ DONE! Files in folder: generated/


STEP 6: Check Results
─────────────────────

macOS & Linux:
ls generated/

Windows (PowerShell):
dir generated/

Windows (Command Prompt):
dir generated/

You should see your converted Excel files in the "generated/" folder


QUICK SUMMARY (For Experienced Users)
──────────────────────────────────────

1. python3 --version               (verify Python 3.7+ installed)
2. pip install -r requirements.txt (install dependencies)
3. python3 excel_generator.py      (run the script)
4. ls generated/                   (or dir generated/ on Windows)


FILES YOU NEED
──────────────

REQUIRED (minimum):
- excel_generator.py      Main Python script
- requirements.txt        Dependencies list

OPTIONAL (documentation):
- Start Here.md          Full English documentation

TROUBLESHOOTING
───────────────

Problem: "command not found: python3"
Solution: Install Python from https://www.python.org/downloads/

Problem: "ModuleNotFoundError: No module named 'docx'"
Solution: Run: pip install -r requirements.txt

Problem: "No .docx files found with 'wydarzenia' in name"
Solution: 
- Check filename contains "wydarzenia"
- Extension must be .docx (not .doc)
- File must be in the same directory as the script


WHAT THE SCRIPT DOES
────────────────────

1. Searches current folder for files with "wydarzenia" in the name
2. Opens each Word (.docx) file
3. Extracts: Dates, Priorities, Client names
4. Creates Excel file with formatted data
5. Saves to "generated/" folder automatically
6. Shows progress messages for each step


PROJECT STRUCTURE
─────────────────

After setup, your folder should look like:

excel-converter/
├── excel_generator.py          Main script
├── requirements.txt            Dependencies
├── wydarzenia_05.25b.docx      Your Word files
├── (other .docx files)
└── generated/                  Created automatically
    ├── maj25.xlsx
    ├── czerwiec25.xlsx
    └── (more Excel files)


FILE FORMATS
────────────

INPUT (Word documents):
Format: .docx (Microsoft Word Open XML Format)
Content: Structured event data with fields:
  - Date: DD Month YYYY HH:MM
  - Priority: Critical, Elevated, High, etc.
  - Client: Client/Company name

OUTPUT (Excel spreadsheets):
Format: .xlsx (Microsoft Excel)
Structure:
  Row 1: Month name (header)
  Row 5: Column headers (No., Event Type, Date, Time, Client)
  Row 6+: Event data


REQUIREMENTS
────────────

python-docx==0.8.11    (Read Word documents)
openpyxl==3.10.1       (Create Excel files)

Both installed automatically with:
pip install -r requirements.txt


🔧 FUNCTIONS & TECHNIQUES USED
────────────────────

1. Document Reading (python-docx)
Opens .docx files and extracts all text paragraphs automatically
2. Pattern Matching (Regex)
Searches text for "Date:", "Priority:", "Client:" patterns
3. Data Parsing
Extracts dates (DD Month YYYY HH:MM), priorities, client names
4. Format Conversion
Transforms: May → 05, "Critical" → priority type, dates → DD.MM.YYYY
5. Excel Creation (openpyxl)
Generates .xlsx files with formatted cells, headers, proper widths
6. Batch Processing
Processes multiple Word files in single run, creates separate Excel per file
7. Error Handling
Validates data completeness, handles missing fields, continues on errors
8. Month Detection
Analyzes filename to determine month and names Excel file accordingly
9. Encoding Support
Handles UTF-8 Polish characters (ą, ć, ę, ł, ń, ó, ś, ź, ż)
10. Automatic Folder Management
Creates "generated/" directory if missing, organizes output files

SUPPORT
───────

For questions or issues:

1. Excel Report Generator.md (English documentation)
2. Check code comments in excel_generator.py

VERSION INFO
────────────

Script: EXCEL GENERATOR v1.0
Language: Python 3.7+
Edition: Python Edition
Status: Production Ready
Created by: Mariusz Grzelak


NEXT STEPS
──────────

1. Follow STEP 1-6 above
2. Add your Word files to the folder
3. Run: python3 excel_generator.py
4. Check "generated/" folder for results

That's it! Your Excel files are ready to use.

