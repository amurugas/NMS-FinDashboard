# Hotel Financial Dashboard

A Streamlit dashboard for hotel finances, backed by Google Sheets.

## Features
- Revenue, expense, net profit, and margin KPIs
- Revenue vs expense monthly comparison
- Net profit trend
- Expense mix pie chart and category bar chart
- Expense filtering by month and category
- Download filtered data to Excel

## Expected Google Sheets tabs

### 1) MonthlySummary
Required columns:
- Month
- Total Revenue
- Total Expenses
- Net Profit

Optional:
- % Increase

### 2) Expenses
Required columns:
- Date
- Category
- Description
- Amount

Optional:
- Payment Mode
- Remarks

## Local run
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Streamlit secrets
Create `.streamlit/secrets.toml` locally.

### Private Google Sheet
```toml
data_source = "private_google_sheet"
google_sheet_key = "YOUR_GOOGLE_SHEET_KEY"
monthly_sheet_name = "MonthlySummary"
expenses_sheet_name = "Expenses"

[gcp_service_account]
type = "service_account"
project_id = "..."
private_key_id = "..."
private_key = "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
client_email = "..."
client_id = "..."
token_uri = "https://oauth2.googleapis.com/token"
auth_uri = "https://accounts.google.com/o/oauth2/auth"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "..."
universe_domain = "googleapis.com"
```

### Public CSV exports
```toml
data_source = "public_csv"
monthly_sheet_name = "https://docs.google.com/spreadsheets/d/.../export?format=csv&gid=..."
expenses_sheet_name = "https://docs.google.com/spreadsheets/d/.../export?format=csv&gid=..."
```

## Deploy on Streamlit Community Cloud
1. Push `app.py`, `requirements.txt`, and `.streamlit/config.toml` to GitHub.
2. Do **not** commit `secrets.toml`.
3. In Streamlit Community Cloud, click **Create app** and select the repo and `app.py` as the entrypoint.
4. Add your secrets in the app settings using the Secrets editor.
5. Deploy.

## Notes
- Share the Google Sheet with the service account email if you use the private-sheet option.
- If the sheet layout changes, update the tab names or column headers.
