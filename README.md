# Excel Essentials Pack

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

**Excel Essentials Pack** is a powerful Visual Studio Tools for Office (VSTO) add-in designed to streamline everyday Excel tasks, enhance data manipulation capabilities, and provide robust macro and workbook management tools. Whether you are cleaning data, importing large text files, or managing VBA modules, this add-in acts as a Swiss Army knife for Microsoft Excel power users.

---

## 🚀 Key Features

Based on the core Ribbon operations, the Excel Essentials Pack offers a wide array of utilities, grouped into the following categories:

### 🧹 Data Cleaning & Manipulation
* **Rapid Type Conversion:** Instantly convert selected ranges to text or values, evaluate formulas in-place, and repaste formulas as values.
* **Cell Scrubbing:** Quickly remove empty cells, error values (`#ERR`), and `#N/A` from your selection.
* **String Operations:** Bulk prepend or append text to selected cells, and trim leading/trailing spaces across massive datasets.
* **Duplicate Management:** Advanced tools to remove duplicates across whole ranges or column-by-column.
* **Fill & Formatting:** Auto-fill empty cells with the value from above, format strings to dates, and instantly clear formatting or clear data outside of the current selected region.

### 🔍 Filtering, Searching, & Visuals
* **Advanced Filtering:** Filter columns by range, exclude by range, or filter using Regular Expressions (Regex). Includes a quick "Flip Filter" toggle.
* **Dynamic Coloring:** Automatically colorize cells or rows based on unique values or specific text matches for quick visual data parsing.
* **Visibility Control:** Instantly hide rows or columns containing specific text, and permanently delete hidden rows or columns to clean up sheets.
* **Search & Select:** Custom search dialogs for headers/columns, and utilities to easily select current regions while ignoring header rows.

### 📁 Import, Export, & File Management
* **Smart Importing:** Robust tools to import text files (even those with bad records or legacy formats). Seamlessly handles large text files by loading them directly into the Excel Data Model.
* **Bulk Exporting:** Save active worksheets or all worksheets as individual `.xlsx` or delimited `.txt` files.
* **Table Splitting:** Divide large tables into designated parts and save them automatically.
* **Workbook/Sheet Utilities:** Duplicate workbooks, move/copy selected worksheets to new workbooks, batch rename selected worksheets, or safely delete unused "other" worksheets.

### ⚙️ Macro & VBA Management
* **Centralized Macro Execution:** Run predefined macros from a centralized `MacrosWbName` or `FunctionsWbName` workbook.
* **VBA Updating & Exporting:** Tools to update `.bas` modules and `.macro` files directly from the ribbon. Easily export all macros from a selected workbook or create update files from active VBA code.
* **Mapping Verification:** Built-in XML parsing to check Ribbon Button ID to VBA Subroutine mappings.

---

## 🛠️ Dependencies

To build, run, or install the Excel Essentials Pack, the following dependencies are required:

* **Microsoft Excel:** Designed for modern desktop versions of Microsoft Excel (Office 365 / Excel 2016+ recommended).
* **.NET Framework:** Required for the VSTO runtime environment (check specific project properties for the exact target framework, typically .NET Framework 4.7.2 or .NET Standard/Core equivalents).
* **Visual Studio Tools for Office (VSTO) Runtime:** Must be installed on the host machine to execute the add-in.
* **Interop Assemblies:** Uses `Microsoft.Office.Interop.Excel` for COM object manipulation.

---

## 📄 License
**MIT License**

Copyright (c) 2026 abc17tg
