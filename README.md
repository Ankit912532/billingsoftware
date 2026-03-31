# GST Billing Software

A Python/Tkinter desktop billing application for Indian GST invoicing with PDF generation, Excel ledger, and optional Google Sheets sync.

## Features
- Generate GST-compliant quotation/invoice PDFs
- Auto-save all bills to a local Excel ledger (`all_bills.xlsx`)
- Company logo support in PDF header
- Optional live sync to Google Sheets
- Editable company settings via GUI

## Requirements
```bash
pip install openpyxl pillow reportlab gspread google-auth
```

## Setup

1. **Clone the repo**
2. **Add your logo** — place `logo.jpg` in the same folder as `billing_app.py`
3. **Configure your company** — edit `company_config.json` with your details, or use the *Company Settings* tab in the app
4. **Run the app**
   ```bash
   python billing_app.py
   ```

## company_config.json Fields

| Field | Description |
|---|---|
| `company_name` | Your company/trade name |
| `legal_name` | Proprietor's legal name |
| `gstin` | Your GSTIN number |
| `address`, `phone`, `email` | Contact details |
| `bank_name`, `account_no`, `ifsc`, `branch` | Bank details for invoice footer |
| `cgst_rate`, `sgst_rate` | Default GST rates (e.g. `"2.5"`) |
| `logo_path` | Path to logo file (default: `logo.jpg` in same folder) |
| `gsheet_enabled` | `true` to enable Google Sheets sync |
| `gsheet_id` | Your Google Sheet ID |

## Google Sheets Sync (Optional)
Follow the in-app instructions under the **Google Sheets** tab to connect a service account.

## Files
| File | Purpose |
|---|---|
| `billing_app.py` | Main application |
| `company_config.json` | Your company configuration |
| `logo.jpg` | Company logo (add your own) |
| `all_bills.xlsx` | Auto-generated bills ledger (gitignored) |
| `google_credentials.json` | Google service account key (gitignored, add your own) |

## License
MIT
