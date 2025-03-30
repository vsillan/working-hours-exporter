# Working Hours Exporter

A personal helper tool.

A Python script that exports working hours from Google Sheets to a PDF report and uploads it to Google Drive.

## Setup

1. Create a Google Cloud Project and enable the Google Sheets API and Google Drive API
2. Create OAuth 2.0 credentials and download them as `credentials.json`
3. Install Python requirements:

```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
pip install -r requirements.txt
```

4. Set up your configuration:
   - Copy `config.template.json` to `config.json`
   - Edit `config.json` with your values:
   ```json
   {
     "spreadsheet_id": "your_spreadsheet_id_here",
     "sheet_name": "Sheet1",
     "output_pdf": "working_hours.pdf",
     "drive_folder_id": "your_drive_folder_id_here"
   }
   ```
   - `spreadsheet_id`: The ID from your Google Sheets URL
   - `sheet_name`: The name of the sheet tab (e.g., "3/2025")
   - `output_pdf`: The name of the PDF file to generate
   - `drive_folder_id`: The Google Drive folder ID where the PDF will be uploaded

## Usage

Run the script:

```bash
python working_hours_exporter.py
```

The script will:

1. Read data from your Google Sheet
2. Generate a PDF report
3. Upload the PDF to Google Drive

## Notes

- On first run, you'll need to authorize the application through your Google account
- The authorization token will be saved in `token.pickle` for future use
- `config.json` contains sensitive information and is ignored by git
