# dLocal Pending Debits Web Application

A modern web application for automating dLocal pending debits processing and journal entry generation.

## Features

- 🎯 **Drag & Drop Interface**: Simply drop your Excel file or click to browse
- 📊 **Automatic Processing**: Detects file month and filters for following month transactions
- 💰 **Smart Calculations**: Formula-driven net amount calculations (Debit - Return)
- 📝 **Journal Entry Generation**: Automatic JE creation with proper debit/credit allocation
- 📥 **Dual Download Options**:
  - Full Workbook: All transactions + Summary + Journal Entry
  - JE Template Only: Ready-to-import journal entry

## Installation

1. **Install Python Dependencies**:
```bash
pip install -r requirements.txt
```

2. **Run the Application**:
```bash
python app.py
```

3. **Access the App**:
Open your browser and navigate to: `http://localhost:5000`

## Usage

### Step 1: Upload File
- Drag and drop your transaction Excel file (e.g., "09 Control Gusto Inc Septiembre 2025.xlsx")
- Or click the upload box to browse for your file
- Click "Process File"

### Step 2: Review Results
The results page shows:
- **File Period**: The month from your uploaded file
- **Filtered For**: The following month (automatically calculated)
- **Transaction Count**: Number of transactions in the following month
- **Calculation Summary**: 
  - Total ACH Debit Amount
  - Total ACH Return Amount
  - Net Amount (highlighted in green/red)
- **Journal Entry Preview**: Shows the generated JE with proper accounts

### Step 3: Download
Choose your download option:
- **Full Workbook**: Contains all transaction lines with highlights, summary, and JE
- **JE Template Only**: Just the journal entry for direct import

## File Format Requirements

Your input file should contain these columns:
- `Date`
- `ACH_DEBIT_AMOUNT`
- `ACH_RETURN_AMOUNT`
- `Date processed`
- `CN` (optional)
- `DN` (optional)

## Journal Entry Logic

**For Positive Net Amount:**
- Debit: 22010 - Customer Funds Obligation : Customer Funds Liability
- Credit: 21017 - Other Current Liabilities : Accrued Liabilities - Platform

**For Negative Net Amount:**
- Credit: 22010 - Customer Funds Obligation : Customer Funds Liability
- Debit: 21017 - Other Current Liabilities : Accrued Liabilities - Platform

## Technical Details

- **Framework**: Flask 3.0
- **Excel Processing**: openpyxl 3.1.2
- **Port**: 5000 (default)
- **Max File Size**: 16MB

## Folder Structure

```
dlocal pending/
├── app.py                 # Main Flask application
├── requirements.txt       # Python dependencies
├── templates/
│   ├── index.html        # Upload page
│   └── result.html       # Results page
├── static/
│   ├── css/
│   │   └── style.css     # Styling
│   └── js/
│       └── main.js       # Drag & drop functionality
├── uploads/              # Temporary upload folder
└── outputs/              # Generated output files
```

## Support

For issues or questions, please refer to the original automation script documentation.
