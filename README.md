# OPSKE Automation Agent

## Overview
This project provides a full automation agent (Python + Playwright + Tkinter GUI) for uploading required supporting documents to the **OPSKE** government portal.  
It validates both **Excel entries** and **local files** before automating the upload and submission process.

---

# About OPSKE (Greek State Aid Management System)
**OPSKE** (Integrated Information System for State Aid Management) is the official digital platform used by the Greek government to manage **EU‑funded investment programs**.

Applicants must upload compliance documents through the OPSKE web interface.  
Since the platform provides **no API access**, all tasks (validation, file selection, date picking, uploading, saving, submitting) normally must be done *manually*.

This automation agent replaces that manual workflow with a structured, error‑proof process.

---

# ✔ What the Agent Actually Checks (Excel + Files)

## 1. **Excel Validation Layer**
Before any upload occurs, the agent performs a *full validation pass* on the Excel file.  
If anything is invalid, the process cannot start — ensuring OPSKE‑ready data.

The following checks are performed:

### Required Excel Columns
The Excel file must contain the following columns (original Greek OPSKE names + English meaning):

- `Επωνυμία – ΑΦΜ` *(Company Name & VAT Number)*  
  Format: "Company Name – #########" and the VAT number must pass the official Greek AFM checksum.

- `Κωδικός έργου` *(Project Code)*  
  Unique identifier of the funded project (e.g., PROJECTNAME-12345).

- `Κωδικός Δικαιολογητικού` *(Document Type Code)*  
  Two-level code in the format NN.NN (e.g., 01.03).

- `Όνομα αρχείου` *(File Basename)*  
  Expected filename (without extension) used to match the actual file in the selected folder.

- `Ημερομηνία έκδοσης δικαιολογητικού` *(Document Issue Date)*  
  Must be a valid date and not in the future.

- `Ημερομηνία λήξης δικαιολογητικού` *(Document Expiration Date)*  
  Must be a valid date and not in the past.

- `Παρατηρήσεις ΟΠΣΚΕ` *(OPSKE Notes / Comments)*  
  Optional free-text notes inserted into the OPSKE portal during upload.


If any are missing → **hard error**.

### **B. AFM Validation**
- Confirms format: `"Company Name – #########"`
- Applies the official **AFM checksum algorithm**  
- Invalid AFM = **Excel error**

### **C. Project Code Validation**
Checks format:  
```
<letters/numbers/Greek>-<number>
```

### **D. Document Code Validation**
Must match:  
```
NN.NN
```

### **E. Issue / Expiry Date Validation**
- Must be valid dates  
- Issue date cannot be in the future  
- Expiry date cannot be in the past  

### **F. Automatic Error Reporting in GUI**
All Excel errors are shown in the GUI with:
- Red entries  
- Exact row and column causing the failure  
- Clear explanation (“Invalid AFM”, “Wrong date format”, etc.)

Only if **zero Excel errors** remain does the agent allow the upload phase.

---

## 2. **File-System Validation Layer**
For each row in the Excel file, the agent performs strict checks on the corresponding local file.

### **A. File Presence Check**
The agent searches the selected folder for a file whose **basename** matches the Excel field `Όνομα αρχείου`.

If no file is found → **error**, row is skipped.

### **B. Extension Check**
Allowed:
- `.pdf`
- `.jpg`
- `.jpeg`
- `.png`

If not in allowed list → **error**

### **C. File Size Check**
Maximum allowed size: **10 MB**

If exceeded → **error**

Again, all errors appear live in the GUI before uploads begin.

Only files that pass the Excel validation **and** file validation proceed to the automation stage.

---

# ✔ Automation Features
Once checks are passed, the agent:

- Logs into OPSKE via Playwright
- Navigates to the Supporting Documents section
- Selects:
  - Beneficiary (Επωνυμία – ΑΦΜ)
  - Project Code
  - Document Type
- Fills notes
- Selects dates using the **calendar widget** (not raw typing)
- Uploads the validated file
- Performs either:
  - **Save**  
  - **Submit**
- Updates the Excel file:
  - `Αποθήκευση`
  - `Υποβολή`
  - `Ημ/νία & ώρα υποβολής`

A live Tkinter GUI shows:
- Current file upload progress
- Total progress bar
- A results table (green = OK, red = fail)
- A summary of skipped or previously submitted documents

---

# Installation
```bash
pip install -r requirements.txt
```

# Running the Application
```bash
python opske_automatation_agent.py
```

---

# Notes
- You must set your OPSKE credentials inside the script (`USERNAME`, `PASSWORD`)
- Excel must strictly follow the required column structure
- Do not modify the date fields manually unless consistent with OPSKE formats
- Ensure stable internet connection during uploads

---

# Disclaimer
This tool automates interaction with an official government portal.  
Use responsibly and in compliance with all legal, ethical, and organizational requirements.
