# Excel Reconciliation App – Books vs GST Portal

A Python-based web application built with **Streamlit** to reconcile purchase data from an accounting Books Excel file against data downloaded from the GST Portal.

The application helps identify matching invoices, mismatches, missing invoices, and invoices that require manual review.

## Features

- **Upload Excel files:** Upload both Books and GST Portal Excel files (`.xlsx`).
- **Automatic header detection:** Detects the actual table header even when the Excel file contains report titles or blank rows above it.
- **Flexible column mapping:** Select which Books columns should be compared against the corresponding GST Portal columns.
- **GSTIN and invoice matching:** Matches invoices primarily using GSTIN and normalized invoice numbers.
- **Invoice number normalization:** Handles common differences in invoice number formats, such as prefixes, separators, and financial-year references.
- **GSTIN mismatch identification:** Flags invoices with a matching invoice number but a different GSTIN for manual review.
- **Mandatory field comparison:** Includes the GST Portal fields marked with an asterisk (`*`).
- **Optional field comparison:** Allows additional Portal fields to be selected.
- **Tax comparison:** Calculates and compares total tax from IGST, CGST, SGST and Cess.
- **Configurable tolerance:** Set an acceptable amount difference tolerance.
- **Reconciliation summary:** View invoice counts by reconciliation status.
- **Status filtering:** Filter the results to focus on specific reconciliation statuses.
- **Excel report download:** Download the complete reconciliation results and summary in an Excel workbook.

## Reconciliation Statuses

| Status | Meaning |
|---|---|
| Exact Match | GSTIN and invoice number match, and all selected comparison fields are within the configured tolerance. |
| Mismatch | GSTIN and invoice number match, but one or more selected comparison fields differ. |
| Needs Review | Invoice number matches, but GSTIN does not match exactly. Manual verification is required. |
| Ambiguous Match | Multiple possible matching invoices were found. Manual verification is required. |
| Only in Books | Invoice is present in Books but no matching invoice was found in the Portal data. |
| Only on Portal | Invoice is present in Portal data but no matching invoice was found in Books. |

## Technology Used

- Python
- Streamlit
- Pandas
- OpenPyXL

## Requirements

Before running the application, install:

1. Python
2. Visual Studio Code (recommended)
3. Git (optional, if downloading through Git)

The application is designed to run locally on your computer.

## Download the Project

### Option 1: Download ZIP from GitHub

1. Open this project's GitHub repository.
2. Click the green **Code** button.
3. Select **Download ZIP**.
4. Extract the downloaded ZIP file to a folder on your computer.
5. Open the extracted project folder.

### Option 2: Clone using Git

Open Terminal and run:

```bash
git clone YOUR_GITHUB_REPOSITORY_URL
```

Replace `YOUR_GITHUB_REPOSITORY_URL` with the actual URL of this repository.

Then navigate into the downloaded project folder:

```bash
cd ExcelReconcile
```

## Installation Instructions

### Step 1: Open the project folder

Open the downloaded project folder in Visual Studio Code.

### Step 2: Create a virtual environment

Run:

```bash
python3 -m venv .venv
```

### Step 3: Activate the virtual environment

**Mac / Linux:**

```bash
source .venv/bin/activate
```

**Windows (Command Prompt):**

```cmd
.venv\Scripts\activate.bat
```

**Windows (PowerShell):**

```powershell
.venv\Scripts\Activate.ps1
```

### Step 4: Install required libraries

Run:

```bash
pip install streamlit pandas openpyxl
```

## Run the Application

Once installation is complete, run:

```bash
streamlit run app.py
```

The application will open in your default web browser.

If it does not open automatically, copy the local URL displayed in Terminal (usually `http://localhost:8501`) and open it in your browser.

## How to Use the Application

### Step 1: Upload Excel files

- Upload your Books Excel file.
- Upload your GST Portal Excel file.
- Both files must be in `.xlsx` format.

### Step 2: Select matching columns

Choose the appropriate columns for:

- Books GSTIN
- Portal GSTIN
- Books invoice number
- Portal invoice number

These columns are used to identify corresponding invoices.

### Step 3: Map comparison columns

Select which Books column should be compared with each Portal field.

The application automatically identifies the expected mandatory Portal fields where possible.

You can also select optional Portal fields for additional comparisons.

If a Books column has no corresponding Portal field, leave it unselected.

### Step 4: Set amount tolerance

Specify the acceptable difference in amounts.

For example, a tolerance of `0.01` allows differences up to ₹0.01 without treating them as a mismatch for numeric comparisons.

### Step 5: Reconcile files

Click **Reconcile Files** and wait for the reconciliation process to complete.

### Step 6: Review the results

Review:

- Total invoices
- Exact matches
- Mismatches
- Invoices only in Books
- Invoices only on Portal
- Invoices requiring manual review

Use the status filter to view specific categories.

### Step 7: Download the report

Click **Download Reconciliation Report** to download the Excel output.

The downloaded workbook contains:

- **Reconciliation:** Detailed invoice-level reconciliation results.
- **Summary:** Count of invoices under each reconciliation status.

## Important Notes

- Keep the original Excel files unchanged and maintain a backup before reconciliation.
- Check the column mapping carefully before running the reconciliation.
- A GSTIN mismatch is flagged for review rather than automatically treated as an exact match.
- Invoice number normalization may simplify invoice numbers for matching. Always review ambiguous matches and unusual invoice formats.
- The application compares only the fields selected in the mapping section.
- The configured amount tolerance applies to numeric field comparisons.
- The application is a reconciliation aid. Review the output and verify exceptions before making accounting or GST-related decisions.

## Privacy

The application runs locally on the computer where it is started. Uploaded Excel files are processed by the local Streamlit application; this project does not include a feature for uploading the files to a separate hosted server.

Do not share your Excel files or generated reconciliation reports publicly, as they may contain confidential business information.

## Project Structure

```text
ExcelReconcile/
│
├── app.py
├── README.md
└── .venv/              # Local virtual environment (created during setup)
```

## Disclaimer

This application is intended to assist with Excel-based reconciliation. It does not replace professional accounting review, GST compliance checks, or verification against the original records.

