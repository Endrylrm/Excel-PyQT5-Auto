# Excel XLS Automation APP

Built for automation of Excel files, to get only the data that is relevant to you.

## Features

As of right now, those are the features:

1. Export selected columns with a optional sorting.
2. Concatenations of files and their sheets.

## Dependencies

The Excel Automation App Dependencies:

1. Pandas
2. PyQT5 (will update it to PyQT6, when PyInstaller have support for PyQT6/PySide6)

## Usage

1. Get Excel XLS.
2. Load Excel file you want to get the data.
3. Use the features you might need right now.
4. Export to Excel.

## Building

If you want to package/build into a executable file, use PyInstaller.

<details>
<summary>Windows / Mac OS</summary>
<br>

> PyInstaller --onefile -w main.py -i Excel-automation-icon.ico

</details>

<details>
<summary>Linux</summary>
<br>

> PyInstaller --onefile main.py -i Excel-automation-icon.ico

</details>

