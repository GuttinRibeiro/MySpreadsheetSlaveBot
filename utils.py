from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient import discovery
import numpy as np

class GoogleAPIHandler:
    def __generate_spreadsheet(self, scope, filename="creds.json", api='sheets', version='v4'):
        creds = ServiceAccountCredentials.from_json_keyfile_name(filename, scope)
        return discovery.build(api, version, credentials=creds).spreadsheets()

    def __init__(self, scope, sheet_id, filename="creds.json", api='sheets', version='v4'):
        self.__spreadsheet = self.__generate_spreadsheet(scope, filename=filename, api=api, version=version)
        self.__spreadsheet_id = sheet_id
        self.__scope = scope
        self.__credential_file = filename

    def get_row(self, sheet, row_number, start, stop):
        # To do: deal with bad inputs
        row_number = row_number + 1
        start = chr(start+97).upper()
        stop = chr(stop+97).upper()
        desired_range = sheet+'!'+start+str(row_number)+':'+stop+str(row_number)
        result = self.__spreadsheet.values().get(spreadsheetId=self.__spreadsheet_id, range=desired_range).execute()
        return result.get('values', [])[0]

    def get_column(self, sheet, column_number, start, stop):
        # To do: deal with bad inputs
        start = start + 1
        stop = stop + 1
        column_number = chr(column_number+97).upper()
        desired_range = sheet+'!'+column_number+str(start)+':'+column_number+str(stop)
        result = self.__spreadsheet.values().get(spreadsheetId=self.__spreadsheet_id, range=desired_range).execute()
        ret = result.get('values', [])
        if not ret:
            return []
        return np.hstack(ret).tolist()

    def get_cell(self, sheet, column, row):
        column = chr(column+97).upper()
        row = row+1
        desired_range = sheet+'!'+column+str(row)
        result = self.__spreadsheet.values().get(spreadsheetId=self.__spreadsheet_id, range=desired_range).execute()
        return result.get('values', [])[0][0]

    def get_table(self, sheet):
        result = self.__spreadsheet.values().get(spreadsheetId=self.__spreadsheet_id, range=sheet).execute()
        return result.get('values', [])

    def write_row(self, sheet, row_number, start, stop, values, isLiteral=True):
        body = {
            'values' : [values]    
        }
        row_number = row_number + 1
        start = chr(start+97).upper()
        stop = chr(stop+97).upper()
        desired_range = sheet+'!'+start+str(row_number)+':'+stop+str(row_number)
        if isLiteral:
            input_opt = "RAW"
        else:
            input_opt = "USER_ENTERED"
        self.__spreadsheet.values().update(spreadsheetId=self.__spreadsheet_id, range=desired_range,
                                          valueInputOption=input_opt, body=body).execute()

    def write_column(self, sheet, column_number, start, stop, values, isLiteral=True):
        body = {
            'values' : values    
        }
        start = start + 1
        stop = stop + 1
        column_number = chr(column_number+97).upper()
        desired_range = sheet+'!'+column_number+str(start)+':'+column_number+str(stop)
        if isLiteral:
            input_opt = "RAW"
        else:
            input_opt = "USER_ENTERED"
        self.__spreadsheet.values().update(spreadsheetId=self.__spreadsheet_id, range=desired_range,
                                          valueInputOption=input_opt, body=body).execute()

    def write_cell(self, sheet, column, row, value, isLiteral=True):
        body = {
            'values' : [[value]]    
        }
        column = chr(column+97).upper()
        row = row+1
        desired_range = sheet+'!'+column+str(row)
        if isLiteral:
            input_opt = "RAW"
        else:
            input_opt = "USER_ENTERED"
        self.__spreadsheet.values().update(spreadsheetId=self.__spreadsheet_id, range=desired_range,
                                          valueInputOption=input_opt, body=body).execute()