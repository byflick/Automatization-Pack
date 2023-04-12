import random
import sys
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


chars = "abcdefghijklmnopqrstuvwxyz1234567890"

def __errors__ (error, type=None):
    if type == 'HttpError':
        if str(error).split('"Requested entity was not found.". Details: ')[1] == '"Requested entity was not found.">':
            sys.exit("Unknown sheet id")
        else:
            sys.exit(error)

class google_sheets():
    def __init__(self, spreadsheets_id='Table id', path='Path to credentials'):
        self.creds = Credentials.from_authorized_user_file(path, ['https://www.googleapis.com/auth/spreadsheets'])
        self.service = build('sheets', 'v4', credentials=self.creds)
        self.sheet = self.service.spreadsheets()
        self.spreadsheets_id = spreadsheets_id

        self.random_name = ''
        for i in range(random.randrange(6, 12)):
            self.random_name += random.choice(chars)

    def find (self, range='Sheet!A1:B2 or Sheet', text='Text for find'):
        if text != 'Text for find':
            pass
        else:
            sys.exit("Enter your text: text='Your text'")

        if range != 'Sheet!A1:B2 or Sheet':
            pass
        elif range == 'Sheet!A1:B2 or Sheet':
            sys.exit("Enter you range: Sheet!A1:B2 or Sheet")
        else:
            range = range.join('A1:Z1000')

        try:
            result = self.sheet.values().get(spreadsheetId=self.spreadsheets_id, range=range).execute()
            values = result.get('values', [])
        except HttpError as error:
            __errors__(error, type='HttpError')

        pos_list = []
        # pos = None
        for row_num, row in enumerate(values):
            for col_num, cell in enumerate(row):
                if cell == text:
                    pos = f"{chr(col_num + 65)}{row_num + 1}"
                    pos_list.append(pos)
                    # break
            # if pos is not None:
            #     break
        if len(pos_list) > 1:
            return pos_list
        else:
            return ''.join(pos_list)

    def info (self, sheet='Sheet name'):
        try:
            if sheet != 'Sheet name':
                spreadsheet_info = self.sheet.get(spreadsheetId=self.spreadsheets_id).execute()
                for sheet_info in spreadsheet_info.get('sheets', []):
                    if sheet_info['properties']['title'] == sheet:
                        return sheet_info
                return None
            else:
                spreadsheet_info = self.service.spreadsheets().get(spreadsheetId=self.spreadsheets_id).execute()
                return spreadsheet_info
        except HttpError as error:
            __errors__(error, type='HttpError')

    def write (self, range='Sheet!A1:B2', text='Text for write'):
        if text != 'Text for write':
            result = self.sheet.values().update(
                spreadsheetId=self.spreadsheets_id, range=range,
                valueInputOption="USER_ENTERED",
                body={"values": [[text]]}).execute()
            return result
        else:
            sys.exit("Enter your text: text='Your text'")

    def append (self, range='Sheet!A1:B2', text='Text for find'):
        if text != 'Text for write':
            result = self.sheet.values().append(
                spreadsheetId=self.spreadsheets_id, range=range,
                valueInputOption="USER_ENTERED",
                body={"values": [[text]]}).execute()
            return result
        else:
            sys.exit("Enter your text: text='Your text'")

    def create_spreadsheet (self, title=None):
        try:
            if title != None:
                name = title
            elif title == None:
                name = self.random_name

            spreadsheet = {
                'properties': {
                    'title': name
                }
            }

            spreadsheet = self.sheet.create(
                body=spreadsheet,
                fields='spreadsheetId').execute()

            spreadsheet = {'title': name,
                           'link': f'https://docs.google.com/spreadsheets/d/{spreadsheet["spreadsheetId"]}',
                           'spreadsheetId': spreadsheet['spreadsheetId']}
            return spreadsheet
        except HttpError as error:
            __errors__(error, type='HttpError')

    def formatting (self, ranges='Sheet!A1:B2', backgroundColor='white', fontColor='white', type='Your formula'):
        color_list = {"red": {'red': 1, 'green': 0.0, 'blue': 0.0},
                      "green": {'red': 0.0, 'green': 1, 'blue': 0.0},
                      "blue": {'red': 0.0, 'green': 0.0, 'blue': 1},
                      "gray": {'red': 0.5, 'green': 0.5, 'blue': 0.5},
                      "yellow": {'red': 1.0, 'green': 1.0, 'blue': 0.0},
                      'purple': {'red': 0.5, 'green': 0.0, 'blue': 0.5},
                      "brown": {'red': 0.6, 'green': 0.4, 'blue': 0.2},
                      "orange": {'red': 1.0, 'green': 0.5, 'blue': 0.0},
                      "black": {'red': 0.0, 'green': 0.0, 'blue': 0.0},
                      "white": {'red': 0.0, 'green': 0.0, 'blue': 0.0},
                      "pink": {'red': 1.0, 'green': 0.75, 'blue': 0.8},

                      "olive": {'red': 0.5, 'green': 0.5, 'blue': 0.0},
                      "silver": {'red': 0.75, 'green': 0.75, 'blue': 0.75},
                      "wheat": {'red': 0.96, 'green': 0.87, 'blue': 0.7},
                      "coral": {'red': 1.0, 'green': 0.5, 'blue': 0.31},
                      "maroon": {'red': 0.5, 'green': 0.0, 'blue': 0.0},
                      "plum": {'red': 0.5, 'green': 0.0, 'blue': 0.5},
                      "teal": {'red': 0.0, 'green': 0.5, 'blue': 0.5},
                      "lime": {'red': 0.0, 'green': 1.0, 'blue': 0.0},
                      "pea": {'red': 0.2, 'green': 0.8, 'blue': 0.2},
                      "cyan": {'red': 0.0, 'green': 1.0, 'blue': 1.0},
                      "royal blue": {'red': 0.25, 'green': 0.41, 'blue': 0.88},
                      "ivory": {"red": 1.0, "green": 1.0, "blue": 0.941},
                      "beige": {"red": 0.961, "green": 0.961, "blue": 0.863},
                      "khaki": {"red": 0.941, "green": 0.902, "blue": 0.549},
                      "golden": {"red": 1.0, "green": 0.843, "blue": 0.0},
                      "coral": {"red": 1.0, "green": 0.5, "blue": 0.314},
                      "salmon": {"red": 0.98, "green": 0.502, "blue": 0.447},
                      "hot pink": {"red": 1.0, "green": 0.412, "blue": 0.706},
                      "fuchsia": {"red": 1.0, "green": 0.0, "blue": 1.0},
                      "lavender": {"red": 0.902, "green": 0.902, "blue": 0.98},
                      "indigo": {"red": 0.294, "green": 0.0, "blue": 0.51},
                      "crimson": {"red": 0.863, "green": 0.078, "blue": 0.235},
                      "charcoal": {"red": 0.212, "green": 0.227, "blue": 0.259},
                      "navy blue": {"red": 0.0, "green": 0.0, "blue": 0.502},
                      "azure": {"red": 0.941, "green": 1.0, "blue": 1.0},
                      "aquamarine": {"red": 0.498, "green": 1.0, "blue": 0.831},
                      "magenta": {"red": 1.0, "green": 0.0, "blue": 1.0}
                      }

        if ranges != 'Sheet!A1:B2':
            pass
        else:
            sys.exit("Enter your range: ranges='Sheet!A1:B2'")

        if type != 'Your formula':
            pass
        else:
            sys.exit("Enter your formatting type: type='Your formula'")


        try:
            sheet_id = self.info(sheet=ranges.split("!")[0])['properties']['sheetId']
            pos_1 = ranges.split(":")[0].split("!")[1]
            pos_2 = ranges.split(":")[1]

            request = {
                'requests': [{
                    'addConditionalFormatRule': {
                        'rule': {
                            'ranges': [{
                                'sheetId': sheet_id,
                                'startRowIndex': int(''.join(i for i in str(pos_1) if i.isdigit()))-1,
                                'endRowIndex': int(''.join(i for i in str(pos_2) if i.isdigit())),
                                'startColumnIndex': int(ord(''.join(i for i in str(pos_1) if i.isalpha()).upper()) - ord('A')),
                                'endColumnIndex': int(ord(''.join(i for i in str(pos_2) if i.isalpha()).upper()) - ord('A')+1),
                            }],
                            'booleanRule': {
                                'condition': {
                                    'type': "CUSTOM_FORMULA",
                                    'values': [{
                                        'userEnteredValue': type
                                    }]
                                },
                                'format': {
                                    'backgroundColor': color_list[backgroundColor.lower()],
                                    'textFormat': {'foregroundColor': color_list[fontColor.lower()]}
                                }
                            }
                        },
                        'index': 0
                    }
                }]
            }

            formatting = self.sheet.batchUpdate(spreadsheetId=self.spreadsheets_id, body=request).execute()
            return formatting
        except HttpError as error:
            __errors__(error, type="HttpError")
