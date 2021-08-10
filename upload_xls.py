#import libraries
import pandas as pd
import os, argparse
import pygsheets

#Read an Input xls workbook
def read_xls_workbook(input_path:str):
    """
    Open an Ms. Excel workbook and reads its sheets
    """
    print(f'Loading Ms. Excel workbook {input_path}')
    wb = pd.ExcelFile(input_path)
    for idx,name in enumerate(wb.sheet_names):
        ws = pd.DataFrame()
        sheet = wb.parse(name)
        ws = ws.append(sheet, ignore_index= True)
        yield idx , name , ws


def authenticate_google_service(credentials_file:str):
    """
    Authenticate to google sheets and google drive services using the JSON credentials file.
    """
    auth = pygsheets.authorize(service_file=credentials_file)
    return auth

def find_google_spreadsheet(auth,spreadsheet_name):
    """
    Search for a google sheet named spreadsheet_name
    """
    found = None
    sheet = None
    try:
        sheet = auth.open(spreadsheet_name)
        #print(f'SpreadSheet {spreadsheet_name} found wih ID:{sheet.id} and URL:{sheet.url}')
        found = True
    except pygsheets.SpreadsheetNotFound as e:
        #print(f'SpreadSheet {spreadsheet_name} not found')
        found = False
    return found,sheet

def create_google_spreadsheet(auth,spreadsheet_name):
    """
    Create a google sheet named spreadsheet_name
    """
    result   = auth.sheet.create(spreadsheet_name)
    sheet_id = result['spreadsheetId']
    sheet    = auth.open_by_key(sheet_id)
    return sheet

def share_google_spreadsheet(auth,sheet,emails):
    """
    Share permissions on google spreadsheet
    """
    #Share with the self and the passed emails to grant write access
    for email in emails:
        sheet.share(email, role='writer',type='user',emailMessage="Attached the uploaded spreadsheet")
    #Share read-only with other users
    sheet.share('',role='reader',type='anyone')


def print_google_spreadsheet_summary(sheet):
    """
    Print a summary of the google spreadsheet
    """
    summary = {
        "ID": sheet.id
      , "Title": sheet.title
      , "URL":sheet.url
      , "Updated":str(sheet.updated)
    }
    # Printing Summary
    print("## Summary ########################################################")
    print("\n".join("{}:{}".format(i, j) for i, j in summary.items()))
    print("###################################################################")

def add_worksheet(sheet,ws,name):
    """
    Copy Ms. Excel data to the spreadsheet
    Create worksheets and copy respective dataframes
    """
    cols_count = len(ws.columns)
    rows_count = len(ws.index)

    sheetfound = False
    for w in sheet.worksheets():
        gws = sheet[w.index]
        if gws.title == name:
           sheetfound = True
           #print("Worksheet Already Exists", gws.id, gws.title, gws.url, gws.rows, gws.cols)

    if not sheetfound:
       gws = sheet.add_worksheet(name, rows=rows_count, cols=cols_count)
    #print(ws, cols_count, rows_count)
    gws.set_dataframe(ws, 'A1')


def is_valid_path(path):
    """
    Validates the path inputted as a parameter and ensures that it is a file path
    """
    if not path:
        raise ValueError(f"Invalid Path")
    if os.path.isfile(path):
       return path
    else:
       raise ValueError(f"Invalid Path {path}")


def parse_args():
    """
    Get user command line parameters
    """
    parser = argparse.ArgumentParser(description="Available Options")

    parser.add_argument('-i'
                       ,'--input_file'
                       ,dest='input_file'
                       ,type=is_valid_path
                       ,required=True
                       ,help = "Enter the path of the file to process")

    parser.add_argument('-c'
                      , '--credentials_file'
                      , dest='credentials_file'
                      , type=str
                      #, required=True
                      , default = ".\\static\\project_credentials.json"
                      , help="Enter the path of the file hosting google service account credentials")

    parser.add_argument('-e'
                      , '--emails'
                      , dest='emails'
                      , type=str
                      #, required=True
                      , default = "bassemmarji@gmail.com"
                      , help="Enter the emails to share with the generated spreadsheet")

    args = vars(parser.parse_args())

    #To Display The Command Line Arguments
    print("## Command Arguments #################################################")
    print("\n".join("{}:{}".format(i,j) for i,j in args.items()))
    print("######################################################################")

    return args

def upload_xls(input_file:str, credentials_file:str,emails:str):
    """
    Upload a local Ms. Excel file to a new google sheet
    """
    wb_name    = os.path.splitext(os.path.basename(input_file))[0]
    print('input Ms. Excel file name = ', wb_name)

    auth = authenticate_google_service(credentials_file=credentials_file)

    while True:
        found, sheet = find_google_spreadsheet(auth, wb_name)
        if found:
            print(f'A worksheet with the same name {wb_name} already exists. This worksheet will be deleted ')
            #print_google_spreadsheet_summary(sheet)
            sheet.delete()
        else:
            break

    sheet = create_google_spreadsheet(auth,wb_name)
    share_google_spreadsheet(auth=auth, sheet=sheet, emails = [emails])
    print_google_spreadsheet_summary(sheet)

    for idx, name, ws in read_xls_workbook(input_file):
        print(f'Reading sheet #{idx}:{name}')
        add_worksheet(sheet, ws, name)


if __name__ == "__main__":
    # Parsing command line arguments entered by user
    args = parse_args()
    upload_xls(input_file       = args['input_file']
              ,credentials_file = args['credentials_file']
              ,emails           = args['emails'])
    #upload_xls(input_file       = ".\\static\\MyTrialSheet.xlsx"
    #           ,credentials_file= ".\\static\\project_credentials.json"
    #           ,emails="bassemmarji@gmail.com")
