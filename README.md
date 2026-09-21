### Automated PDF-to-Excel Data Extraction Pipeline 🚀

An enterprise-ready Python pipeline designed to transform unstructured text from complex, multi-page PDF documents into cleanly normalized, multi-tab Excel workbooks (.xlsx) in seconds. 

### 💼 The Business Problem

Corporate data is often trapped inside unstructured PDF formats—such as financial audits, legal briefs, and presentation transcripts. Manually copy-pasting thousands of rows introduces data-entry typos and eats up dozens of valuable internal company hours. 

This repository provides an **automated alternative**. Using a real-world case study—the *2026 Trinidad and Tobago National Mid-Year Fiscal Review*—this pipeline demonstrates how a messy, multi-page speech transcript can be systematically converted into a formula-ready corporate database in **seconds**. 

### ✨ Core Technical Features

* **Multi-Page Layout Looping:** Seamlessly processes varying structural flows across an unlimited page count without hitting buffer walls.
* **Line-Wrap Stitching:** Advanced algorithmic string buffering that automatically detects and glues together broken multi-line sentences (e.g., long ministry and entity names).
* **Text Normalization Engine:** Employs targeted Regular Expressions (re) and lookup maps to resolve classic PDF spacing anomalies (such as smashed words like andMinistry).
* **Multi-Tab Excel Generation:** Leverages pandas and openpyxl engines to cleanly write isolated data streams into separate, descriptive sheets (Recurrent Expenditure vs. Development Expenditure) with native indexing suppressed.

### 🛠️ The Tech Stack

* **Language:** Python 3.x
* **Core Libraries:** pdfplumber, pandas, openpyxl, re

### 📊 Sample Pipeline Output Structure

When executed, the script flattens fragmented layout arrays into an immaculate, client-ready matrix: 

Head NumberEntityBudget Amount
**01**
President$1.0 million
**03**
Judiciary$39.15 million
**07**
Statutory Authorities Service Commission$285 million
**23**
Office of the Attorney General and Ministry of Legal Affairs$75 million

### 📈 Looking for Data Automation Services?

I build custom, secure, and highly scalable Python extraction tools tailored entirely to your business constraints. Whether you need to process invoices, logistics faxes, legacy archives, or bank statements: 

