# Shopee Payment Reconciliation Automation

**Python automation tool that processes Shopee payment reports, matches order numbers with an internal order book, and automatically updates payment amounts and received dates in Excel.**

---

## 🚀 Overview

This project automates the manual process of reconciling marketplace payments with company order records.  
It reads a payment report (Excel), matches order IDs to the internal order book, and updates the received payment amount and date automatically.  

Key features:

- Reads Shopee payment reports (`Excel 1`) and internal order book (`Excel 2`)  
- Matches order numbers and updates payment amounts and dates  
- Creates a backup of the order book before updating  
- Provides a simple GUI for file selection, status logging, and date management  
- Supports using today’s date or manually entering a payment date  

---

## 🛠️ Tech Stack

- **Python 3.x**  
- **pandas** – for reading and processing Excel files  
- **openpyxl** – for updating Excel workbooks while preserving formatting  
- **tkinter** – for GUI file selection and status logs  
- **threading** – to keep the GUI responsive during processing  
- **shutil & os** – for file backup and management  

---

## ⚙️ How to Use

1. Prepare **Excel 1** (Shopee payment report) with columns: `Order ID`, `Amount`  
2. Prepare **Excel 2** (internal order book) with columns: `Order ID`, `Received Amount`, `Payment Date`  
3. Run the Python script:

```bash
python shopee_payment_reconciliation.py
