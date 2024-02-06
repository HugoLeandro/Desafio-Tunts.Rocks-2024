# Import necessary libraries
import math
import os
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Define constants
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SAMPLE_SPREADSHEET_ID = '1VH6s16fKcYGM2zXL3CnuXgVzGYs17H99cDcWpKC8XZQ'
SAMPLE_RANGE_NAME = 'engenharia_de_software!A3:H27'
TOTAL_CLASSES = 60  # Assuming total number of classes is 60


def calculate_student_status(values):
    """Calculate student status based on attendance and grades."""
    headers = values[0]
    students = values[1:]

    for student in students:
        # Convert string values to integers for 'Faltas', 'P1', 'P2', 'P3'
        student[2:6] = list(map(int, student[2:6]))

        # Calculate average grade
        avg_grade = sum(student[3:6]) / 3

        # Check if student has more than 25% absences
        if student[2] > TOTAL_CLASSES * 0.25:
            student.append("Reprovado por Falta")
            student.append(0)
        else:
            # Check student's situation based on average grade
            if avg_grade < 5:
                student.append("Reprovado por Nota")
                student.append(0)
            elif 5 <= avg_grade < 7:
                student.append("Exame Final")
                # Calculate grade needed for final approval
                naf = math.ceil(2 * 5 - avg_grade)
                student.append(naf)
            else:
                student.append("Aprovado")
                student.append(0)

        print(f"Student {student[1]} situation: {student[6]}, grade for final approval: {student[7]}")

    return [headers] + students


def main():
    """Main function to read and update Google Sheets."""
    creds = None

    # Load credentials from file, if it exists
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)

    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)

        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        # Build the service
        service = build('sheets', 'v4', credentials=creds)

        # Read data from Google Sheets
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                    range=SAMPLE_RANGE_NAME).execute()

        values = result['values']
        updated_values = calculate_student_status(values)

        # Update the Google Sheets with the new values
        sheet.values().update(
            spreadsheetId=SAMPLE_SPREADSHEET_ID,
            range=SAMPLE_RANGE_NAME,
            valueInputOption="USER_ENTERED",
            body={"values": updated_values}
        ).execute()

    except HttpError as err:
        print(err)


if __name__ == '__main__':
    main()