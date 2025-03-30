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
import json

# If modifying these scopes, delete the file token.pickle.
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.file",
]


def load_config():
    """Load configuration from config.json file."""
    config_path = "config.json"
    if not os.path.exists(config_path):
        print(f"Error: {config_path} not found!")
        print(
            "Please copy config.template.json to config.json and fill in your values."
        )
        exit(1)

    try:
        with open(config_path, "r") as f:
            config = json.load(f)
            required_fields = [
                "spreadsheet_id",
                "range_name",
                "output_pdf",
                "drive_folder_id",
            ]
            for field in required_fields:
                if field not in config:
                    print(f"Error: Missing required field '{field}' in config.json")
                    exit(1)
            return config
    except json.JSONDecodeError:
        print("Error: config.json is not valid JSON")
        exit(1)
    except Exception as e:
        print(f"Error reading config.json: {str(e)}")
        exit(1)


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

    # Find the row index of "Invoiceable total"
    invoiceable_total_index = df[df[date_column] == "Invoiceable total"].index[0]

    # Filter out rows after "Invoiceable total"
    df = df.iloc[: invoiceable_total_index + 1]

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

    # Create a custom style for notes that preserves whitespace
    notes_style = ParagraphStyle(
        "Notes",
        parent=styles["Normal"],
        fontSize=10,
        leading=12,  # Line spacing
        spaceBefore=0,
        spaceAfter=0,
        preserveWhitespace=True,
        alignment=0,  # Left alignment
    )

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
        # Convert notes to Paragraph for better wrapping
        notes = (
            row[headers[3]] if len(headers) > 3 and pd.notna(row[headers[3]]) else ""
        )
        # Replace newlines with <br/> tags for proper line breaks
        notes = str(notes).replace("\n", "<br/>")
        notes_paragraph = Paragraph(notes, notes_style)
        table_data.append(
            [
                row[headers[0]],  # Date
                row[headers[1]],  # Day
                row[headers[2]],  # Hours
                notes_paragraph,  # Notes as Paragraph
            ]
        )

    # Create table with specific column widths
    col_widths = [1.2 * inch, 0.5 * inch, 0.8 * inch, 4.5 * inch]  # Adjusted widths
    table = Table(table_data, colWidths=col_widths)
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
                ("TOPPADDING", (0, 0), (-1, -1), 4),  # Added top padding
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),  # Added bottom padding
                ("VALIGN", (0, 0), (-1, -1), "TOP"),  # Align content to top
                ("SPAN", (0, 0), (-1, 0)),  # Make header row span all columns
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
    # Load configuration
    config = load_config()

    # Get data from Google Sheets
    print("Fetching data from Google Sheets...")
    data = get_sheet_data(config["spreadsheet_id"], config["range_name"])

    if not data:
        print("No data found in the specified range.")
        return

    # Create PDF
    print("Generating PDF...")
    create_pdf(data, config["output_pdf"])

    # Upload to Google Drive
    print("Uploading to Google Drive...")
    file_id = upload_to_drive(config["output_pdf"], config["drive_folder_id"])
    print(f"File uploaded successfully! File ID: {file_id}")


if __name__ == "__main__":
    main()
