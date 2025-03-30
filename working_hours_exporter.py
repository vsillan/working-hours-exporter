import os
import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
import pickle
from datetime import datetime

# If modifying these scopes, delete the file token.pickle.
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.file",
]


def get_google_credentials():
    """Get or refresh Google API credentials."""
    creds = None
    if os.path.exists("token.pickle"):
        with open("token.pickle", "rb") as token:
            creds = pickle.load(token)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        with open("token.pickle", "wb") as token:
            pickle.dump(creds, token)

    return creds


def get_sheet_data(spreadsheet_id, range_name):
    """Fetch data from Google Sheets."""
    creds = get_google_credentials()
    service = build("sheets", "v4", credentials=creds)
    sheet = service.spreadsheets()
    result = (
        sheet.values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
    )
    return result.get("values", [])


def create_pdf(data, output_filename):
    """Create a PDF from the sheet data."""
    # Convert data to pandas DataFrame
    # Ensure we have all column names, even if some are empty
    headers = data[0]
    while len(headers) < 4:  # We need at least 4 columns (A, B, C, D)
        headers.append(f"Column_{len(headers)}")

    df = pd.DataFrame(data[1:], columns=headers)

    # Get the first column name (date column)
    date_column = headers[0]

    # Find the last non-empty row before the totals
    last_data_row = df[
        df[date_column].notna()
        & (df[date_column] != "Total hours")
        & (df[date_column] != "Invoiceable total")
    ].index[-1]

    # Get the totals
    total_hours_row = df[df[date_column] == "Total hours"].iloc[0]
    invoiceable_total_row = df[df[date_column] == "Invoiceable total"].iloc[0]

    # Create PDF
    doc = SimpleDocTemplate(output_filename, pagesize=letter)
    elements = []
    styles = getSampleStyleSheet()

    # Add title
    title_style = ParagraphStyle(
        "CustomTitle",
        parent=styles["Heading1"],
        fontSize=16,
        spaceAfter=30,
        alignment=0,  # Left alignment
    )
    elements.append(Paragraph(f"Working Hours Report - {data[0][0]}", title_style))

    # Create table with daily entries
    table_data = [["Date", "Day", "Hours", "Notes"]]
    # Only process rows up to the last data row (before totals)
    for _, row in df.iloc[: last_data_row + 1].iterrows():
        table_data.append(
            [
                row[headers[0]],  # Date
                row[headers[1]],  # Day
                row[headers[2]],  # Hours
                row[headers[3]] if len(headers) > 3 else "",  # Notes
            ]
        )

    # Create table
    table = Table(table_data)
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.grey),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
                ("ALIGN", (0, 0), (-1, -1), "LEFT"),  # Changed to LEFT alignment
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, 0), 12),
                ("BOTTOMPADDING", (0, 0), (-1, 0), 12),
                ("BACKGROUND", (0, 1), (-1, -1), colors.beige),
                ("TEXTCOLOR", (0, 1), (-1, -1), colors.black),
                ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
                ("FONTSIZE", (0, 1), (-1, -1), 10),
                ("GRID", (0, 0), (-1, -1), 1, colors.black),
                ("WORDWRAP", (0, 0), (-1, -1), True),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),  # Added left padding
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),  # Added right padding
            ]
        )
    )

    elements.append(table)
    elements.append(Spacer(1, 20))

    # Add totals
    total_style = ParagraphStyle(
        "Total",
        parent=styles["Normal"],
        fontSize=12,
        spaceAfter=10,
        alignment=0,  # Left alignment
    )
    elements.append(
        Paragraph(f"Total Hours: {total_hours_row[headers[2]]}", total_style)
    )
    elements.append(
        Paragraph(
            f"Invoiceable Total: €{invoiceable_total_row[headers[2]]}", total_style
        )
    )

    doc.build(elements)


def upload_to_drive(file_path, folder_id=None):
    """Upload a file to Google Drive."""
    creds = get_google_credentials()
    service = build("drive", "v3", credentials=creds)

    file_metadata = {
        "name": os.path.basename(file_path),
        "parents": [folder_id] if folder_id else [],
    }

    media = MediaFileUpload(file_path, mimetype="application/pdf", resumable=True)

    file = (
        service.files()
        .create(body=file_metadata, media_body=media, fields="id")
        .execute()
    )

    return file.get("id")


def main():
    # Replace these with your actual values
    SPREADSHEET_ID = "1YH-O4aFr3ii2_mZ5hLDKBqh1uBFMBGdhbhilIkaLByo"
    RANGE_NAME = "3/2025!A1:Z1000"  # Adjust range as needed
    OUTPUT_PDF = "working_hours.pdf"
    DRIVE_FOLDER_ID = (
        "1-evAZRTSLtFDepZSGd88mPsnpb_nHN-b"  # Optional: ID of the folder to upload to
    )

    # Get data from Google Sheets
    print("Fetching data from Google Sheets...")
    data = get_sheet_data(SPREADSHEET_ID, RANGE_NAME)

    if not data:
        print("No data found in the specified range.")
        return

    # Create PDF
    print("Generating PDF...")
    create_pdf(data, OUTPUT_PDF)

    # Upload to Google Drive
    print("Uploading to Google Drive...")
    file_id = upload_to_drive(OUTPUT_PDF, DRIVE_FOLDER_ID)
    print(f"File uploaded successfully! File ID: {file_id}")


if __name__ == "__main__":
    main()
