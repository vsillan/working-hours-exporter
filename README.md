# Working Hours Exporter

This Python script allows you to:

1. Access data from a Google Sheet containing working hours
2. Generate a PDF report from the sheet data
3. Upload the generated PDF back to Google Drive

## Setup Instructions

1. Install the required dependencies:

```bash
pip install -r requirements.txt
```

2. Set up Google Cloud Project and enable APIs:

   - Go to the [Google Cloud Console](https://console.cloud.google.com/)
   - Create a new project or select an existing one
   - Enable the Google Sheets API and Google Drive API
   - Create credentials (OAuth 2.0 Client ID)
   - Download the credentials and save them as `credentials.json` in the project directory

3. Configure the script:
   - Open `working_hours_exporter.py`
   - Replace `your_spreadsheet_id` with your Google Sheet ID (you can find this in the URL of your sheet)
   - Replace `your_folder_id` with the ID of the Google Drive folder where you want to upload the PDF (optional)
   - Adjust the `RANGE_NAME` if needed (default is 'Sheet1!A1:Z1000')

## Usage

Run the script:

```bash
python working_hours_exporter.py
```

The first time you run the script, it will:

1. Open a browser window for Google authentication
2. Ask you to authorize the application
3. Save the credentials in `token.pickle` for future use

The script will then:

1. Fetch data from your specified Google Sheet
2. Generate a PDF report with the data
3. Upload the PDF to Google Drive
4. Print the file ID of the uploaded PDF

## Notes

- The script requires the Google Sheet to have headers in the first row
- The generated PDF will have a table format with alternating colors for better readability
- Make sure you have sufficient permissions to access the Google Sheet and the target Google Drive folder
