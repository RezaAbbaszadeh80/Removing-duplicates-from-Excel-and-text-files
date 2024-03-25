# Removing-duplicates-from-Excel-and-text-files
This program takes an excel or text file from the same notepad and deletes duplicate and similar items and even words inside brackets and parentheses etc.

# Removing duplicates from Excel and text files

## Table of Contents
- [Introduction](#introduction)
- [Installation](#installation)
- [Usage](#usage)
- [Functionality](#functionality)
- [Keyboard Shortcuts](#keyboard-shortcuts)
- [Contributing](#contributing)
- [License](#license)

## Introduction
This Python script provides a graphical user interface (GUI) for removing duplicate rows from Excel and text files. It utilizes Tkinter for the GUI, Pandas for data manipulation, and keyboard for implementing keyboard shortcuts.

## Installation
Before running the script, ensure you have the required packages installed:
```bash
# pip install keyboard
# pip install pandas
# pip install pyinstaller
# pip install xlrd
```

## Usage

Run the script and click the "Upload Excel or Text file" button to select the file you want to process. The script supports Excel files with extensions .xlsx, .xls, or .xlsm, as well as text files (.txt).
Functionality

    Excel File Handling:
        Reads Excel files into Pandas DataFrame.
        Removes duplicate rows.
        Saves the cleaned data to a new Excel file.
    Text File Handling:
        Cleans text from a text file.
        Removes duplicate lines.
        Saves the cleaned text back to the same file.
    Filtering Values:
        Provides an entry widget to filter values based on a custom pattern.

## Keyboard Shortcuts

    Windows:
        Ctrl+Z: Undo changes.
        Ctrl+A: Select all text in the result box.

## Contributing

Contributions are welcome! If you have any suggestions or improvements, feel free to open an issue or create a pull request on GitHub.
