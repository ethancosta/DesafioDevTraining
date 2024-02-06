import os.path
import math

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = "1Yyry0ttqUWbscnQpvY2Lhb5VsY6fFoXfrd8lx9bdSq4"
# Reads the columns I want to work with on google sheets
SAMPLE_RANGE_NAME = "engenharia_de_software!B4:H28"


def main():
    situation = ""
    creds = None

    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                "credentials.json", SCOPES
            )
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    try:
        service = build("sheets", "v4", credentials=creds)

        # Call the Sheets API
        sheet = service.spreadsheets()
        result = (
            sheet.values()
                .get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=SAMPLE_RANGE_NAME)
                .execute()
        )
        values = result.get("values", [])

        append_values = []

        for i, line in enumerate(values):
            p1 = float(line[2]) / 10  # Converts all values of the third column that I selected to work with
            # in SAMPLE_RANGE_NAME to float and divides by 10, storing the result in p1
            p2 = float(line[3]) / 10
            p3 = float(line[4]) / 10
            m = math.ceil((p1 + p2 + p3) / 3)  # Calculates the average of p1, p2, and p3,
            # rounds up using math.ceil, and stores the result in m
            fouls = int(line[1])  # Converts the second column that I selected to work with
            # in SAMPLE_RANGE_NAME to int and stores the result in fouls
            if m < 5:
                situation = "Reprovado por Nota"
            elif 5 <= m < 7:
                situation = "Exame Final"
            elif m >= 7:
                situation = "Aprovado"
            if fouls >= 15:
                situation = "Reprovado por Falta"

            naf = 0
            if situation == "Exame Final":
                naf = 2 * 5 - m
            append_values.append([situation, naf])

        result = (  # Updates the Google Spreadsheet with the values in append_values
            sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range="G4:H27",
                                  valueInputOption="USER_ENTERED", body={'values': append_values}).execute()
        )

    except HttpError as err:
        print(err)


if __name__ == "__main__":
    main()
