# 📊 a3eco Invoice PDF Builder

[![VBA](https://img.shields.io/badge/Language-VBA%20%2F%20Macro-green?logo=microsoft-excel&logoColor=white)](#)
[![Automation](https://img.shields.io/badge/Focus-Process%20Automation-blue)](#)

> **Stable Version Download:** You can download the ready-to-use Excel workbook from the [Releases](https://github.com/samuelcedillasanchez/a3eco-InvoicePDFBuilder/releases) section.

An Excel utility powered by VBA (Visual Basic for Applications) designed to automatically extract scanned invoice files embedded within the a3eco ledger and compile them into a single, chronologically sorted PDF document according to accounting records.

---

##  Key Features

* ** Automated PDF Consolidation:** Automatically identifies and merges embedded source documents (scanned invoices) into a standardized bookkeeping ledger package.
* ** Format Normalization:** Built-in routine to convert Legacy image formats (`.tif`) into standard web-ready documents (`.pdf`).
* ** Dynamic File Renaming:** Automatically remaps and overwrites filename strings based on the structural criteria of the a3eco billing logs.
* ** Operational Efficiency:** Built-in modular control pipeline allowing complete execution with a single macro trigger or a 3-step structured manual process.

##  Prerequisites

* **PDFtk Server:** This tool requires the **PDFtk** binary utility installed and configured in the system environment path to execute the command-line PDF compilation routines.

##  How to Use

### 1. Environment Preparation
1. Export your billing backup files directly from the **a3eco** system.
2. Move the generated `FACTURAS` folder and this Excel automation file into the local deployment directory path: `\A3\A3ECO\EMP`
3. Duplicate and paste both resources into the regulatory submission directory: `REQAEAT`

### 2. Automated Pipeline Execution
1. Open the Excel workbook and **enable Macros**.
2. The program will automatically run background data parsing routines to audit embedded files, resolve matching references, and invoke PDFtk to compile a consolidated, sorted master PDF document.

###  Manual Step-by-Step Control (Optional)
If your accounting workspace manages multiple business activities or you prefer a granular execution, you can trigger the 3 structural action buttons available in the layout sequence:
1. **Convert TIF to PDF:** Launches the asynchronous image transformation pipeline.
2. **Replace Filenames:** Renames every single PDF file mapping the reference data extracted from the accounting ledger columns.
3. **Merge PDFs in Order:** Invokes the PDFtk backend to consolidate all individual documents while strictly respecting the indexing sequence.

>  *To apply this workflow to alternative scopes, simply update the lookup criteria rows inside the Excel sheet data structure and run the control buttons sequentially.*

##  Performance & Scalability Constraints

* This utility architecture is highly optimized for business records containing **fewer than 3,000 invoices** to prevent local Excel runtime bottlenecks or thread suspension.
* **Scale Up Exception:** If the targeted enterprise logging registers scale past the 3,000 limit, you must manually extend down the calculation formula cells mapped in columns **N, O, and P**.

---

*Designed and coded with 💻 by Samuel Cedilla Sánchez.*
