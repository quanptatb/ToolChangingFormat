---
name: bom_formatter
description: Guidelines and layout documentation for maintaining and extending the BOM Formatter application, including adding new format types, editing spreadsheet styling, and packaging with PyInstaller.
---

# BOM Formatter Codebase Skill

This skill explains the directory structure, design patterns, and development guidelines for the BOM Formatter codebase. Use this as a reference whenever modifying or extending formatting behaviors.

## Directory Structure

The project has been modularized to separate the CLI/Web runner logic from the core formatting components:

- [main.py](file:///c:/Users/Cua/Desktop/proj/ToolChangingFormat/main.py): Entry point for the CLI and Web Server. Handles server endpoints, multipart file upload parsing, HTML page templates, and routing requests to the appropriate formatter.
- `components/`: Core logic package containing formatting functions.
  - [components/__init__.py](file:///c:/Users/Cua/Desktop/proj/ToolChangingFormat/components/__init__.py): Package initializer exposing main sheet-processor entry points.
  - [components/common.py](file:///c:/Users/Cua/Desktop/proj/ToolChangingFormat/components/common.py): Shared utility functions (such as normalizers, date extraction, value parsers, workbook saves) and shared constants/mappings (`SHORT_NAMES`, `CUSTOMER_SHORT_NAMES`).
  - [components/format1_a4.py](file:///c:/Users/Cua/Desktop/proj/ToolChangingFormat/components/format1_a4.py): Core processor, sorting, and styling implementation for **Format 1** (A4 Portrait layout, 13 columns).
  - [components/format2_shopping.py](file:///c:/Users/Cua/Desktop/proj/ToolChangingFormat/components/format2_shopping.py): Core processor and compiler logic for **Format 2** (Giấy đi chợ - two-panel landscape layout).
  - [components/approval.py](file:///c:/Users/Cua/Desktop/proj/ToolChangingFormat/components/approval.py): Core processor, merging, page breaking, and styling logic for **Duyệt định mức** (A4 landscape review layout).

## Development Guidelines

### 1. Extending Formatting Constants or Mappings
* If adding new short-name mappings or customer name abbreviations, edit the dictionaries `SHORT_NAMES` or `CUSTOMER_SHORT_NAMES` inside `components/common.py`.

### 2. Modifying Excel Styles / Page Margins
* Open the corresponding layout component (`format1_a4.py`, `format2_shopping.py`, or `approval.py`).
* Locate the helper function starting with `apply_*_base_format` or `configure_*_print`.
* Adjust cell alignments, borders, font weights, row heights, and print parameters (such as `page_setup.paperSize`, margins, or horizontal centering) there.

### 3. Adding a New Layout Format
* Create a new component file inside `components/` (e.g., `components/my_new_layout.py`).
* Implement sheet parsing, data sorting/merging, base formatting, and setup logic.
* Register and export the processor function inside `components/__init__.py`.
* In `main.py`, update `format_workbook_bytes` to check for the new option and forward the worksheet to your new module.
* Update `index_html` inside `main.py` to add your format option in the UI's layout select list.

## Testing & Packaging

### Local Verification
* Run CLI batch mode to test layout exports in the `Đã gộp` folder:
  ```powershell
  python main.py --batch
  ```
* Run the HTTP server:
  ```powershell
  python main.py --no-browser
  ```

### PyInstaller Packaging
* The codebase uses PyInstaller to build a standalone executable that works on target machines without requiring Python.
* PyInstaller traces dependencies automatically from `main.py`. Because all formatters are imported, they will be bundled inside the package.
* Run packaging using the script:
  ```powershell
  .\build_exe.bat
  ```
* The output is placed in `dist/BomFormatterWeb`.
