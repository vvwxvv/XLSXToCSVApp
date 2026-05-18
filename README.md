# XLSX to CSV Converter App

A modern desktop application for converting Excel files (.xlsx, .xls) to CSV
format with advanced options including sheet selection, delimiter customization,
and encoding control.

```
 ___________________________________________________________
|                                                           |
|          XLSX  -->  CSV  Converter  App                   |
|___________________________________________________________|
```

---

## [*] Features

- Multi-Sheet Support   -- Select and convert individual or all sheets
- Delimiter Options     -- Comma, semicolon, tab, or pipe separators
- Encoding Control      -- UTF-8 and system locale encoding support
- Modern UI             -- Drag and drop interface with frameless window design
- Batch Conversion      -- Convert multiple files or entire folders at once
- Auto-Detection        -- Automatic sheet preview before conversion

---

## [!] Requirements

- Windows 10/11
- Python 3.10+  (for source installation)
- 100MB RAM
- 50MB free space

---

## [+] Installation

### Option 1: Executable (Recommended)

    1. Download the latest release
    2. Extract the archive
    3. Run XLSXToCSV.exe

### Option 2: From Source

    git clone https://github.com/vvwxvv/XLSXToCSVApp.git
    cd XLSXToCSVApp
    python -m venv appenv
    appenv\Scripts\activate
    pip install -r requirements.txt
    python main.py

---

## [>] How to Use

    +-----------------------------------------------+
    | Step 1  |  Launch the application              |
    | Step 2  |  Select your .xlsx or .xls file      |
    | Step 3  |  Select your output directory        |
    | Step 4  |  Choose sheet(s) to convert          |
    | Step 5  |  Configure delimiter and encoding    |
    | Step 6  |  Click "Start Converting"            |
    +-----------------------------------------------+

    Available Options:
    [x] Choose delimiter  -- comma, semicolon, tab, or pipe
    [x] Choose encoding   -- UTF-8 recommended
    [x] Include or exclude the header row

---

## [~] Supported Formats

    Input      : .xlsx  |  .xls
    Output     : .csv
    Delimiters : comma (,)  |  semicolon (;)  |  tab (\t)  |  pipe (|)
    Encoding   : UTF-8 (default)  |  system locale fallback
    Headers    : optional -- include or skip first row
    Sheets     : single sheet or all sheets as separate CSV files

---

## [#] Examples

### Single Sheet Conversion

    Input  : report.xlsx   [ Sheet: "Sales" ]
    Output : report_Sales.csv

### All Sheets Conversion

    Input  : workbook.xlsx   [ Sheets: "Jan", "Feb", "Mar" ]
    Output :
       |-- workbook_Jan.csv
       |-- workbook_Feb.csv
       |-- workbook_Mar.csv

### Custom Delimiter

    Input     : data.xlsx
    Delimiter : Tab
    Output    : data.csv  (tab-separated, compatible with legacy systems)

---

## [?] Troubleshooting

    +---------------------------+--------------------------------------------------+
    | Error                     | Fix                                              |
    +---------------------------+--------------------------------------------------+
    | "File not supported"      | Use .xlsx or .xls only                           |
    | "Sheet not found"         | Refresh sheet list after selecting file          |
    | "Encoding error"          | Switch to UTF-8 with BOM for Excel compatibility |
    | "Empty output file"       | Confirm sheet has data and is not hidden         |
    | "IndexError during build" | Add excludes=['scipy'] to XLSXToCSV.spec         |
    +---------------------------+--------------------------------------------------+

---

## [%] Building from Source

    pip install pyinstaller
    pyinstaller XLSXToCSV.spec

    NOTE: If you encounter an IndexError: tuple index out of range during
    the build step, add excludes=['scipy'] to the Analysis() block in your
    spec file. This is a known PyInstaller and scipy incompatibility and
    does not affect runtime behavior.

    a = Analysis(
        ['main.py'],
        ...
        excludes=['scipy'],   # <-- add this line
        ...
    )

---

## [$] License

MIT License -- see LICENSE file for details.

---

```
  ___________________________________
 |                                   |
 |   Made for efficient Excel data   |
 |          management.  (^_^)       |
 |___________________________________|
```
