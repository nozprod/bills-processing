import os
import pytesseract
import argparse
import logging.config
import logging
from pdf2image import convert_from_path
from invoice2data import extract_data
from invoice2data.extract.loader import read_templates
import google.auth
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Parser les arguments de la ligne de commande
parser = argparse.ArgumentParser()
parser.add_argument("--debug", action="store_true",
                    help="activer le mode de débogage")
args = parser.parse_args()

# Configurer le logger
if args.debug:
    logging.basicConfig(level=logging.DEBUG)
else:
    logging.basicConfig(level=logging.INFO)

def read_templates_from_folder(folder):
    result = []
    for path, subdirs, files in os.walk(folder):
        for name in files:
            if name.endswith(('.yml', '.yaml')):
                file_path = os.path.join(path, name)
                result.extend(read_templates(file_path))
                print(f"Template loaded from {file_path}")
    return result

def extract_data_from_invoice(pdf_path):
    print("Extraction du texte à partir du fichier PDF...")
    templates = read_templates_from_folder('./Templates')
    try:
        data = extract_data(pdf_path, templates=templates)
        print(f"Data extracted: {data}")
        return data
    except Exception as e:
        print(f"Error during extraction: {e}")
        return None

def extract_invoice_data(data):
    if not data:
        return None

    invoice_data = {}

    date = data.get('date')
    if date:
        invoice_data['date'] = date.strftime("%d/%m/%Y")

    invoice_data['ht'] = data.get('ht')
    invoice_data['tva_rate'] = data.get('tva_rate')
    invoice_data['tva_amount'] = data.get('tva_amount')
    invoice_data['amount'] = data.get('amount')

    return invoice_data

def authenticate_google_sheets():
    print("Authentification et création du service Google Sheets...")
    creds = None
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    token_path = 'token.json'
    credentials_path = 'credentials.json'

    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(credentials_path, SCOPES)
        creds = flow.run_local_server(port=0)

    return build('sheets', 'v4', credentials=creds)

def update_google_sheet(service, sheet_id, data):
    print("Mise à jour de la feuille de calcul Google Sheets...")
    range_name = 'Sheet1!A1:E1'
    values = [[data['date'], data['ht'], data['tva_rate'], data['tva_amount'], data['ttc']]]
    body = {'values': values}

    try:
        result = service.spreadsheets().values().append(
            spreadsheetId=sheet_id, range=range_name,
            valueInputOption='USER_ENTERED', insertDataOption='INSERT_ROWS', body=body).execute()
        print('{0} cells appended.'.format(result.get('updates').get('updatedCells')))
    except HttpError as error:
        print('An error occurred: {0}'.format(error))
        return None

if __name__ == '__main__':
    pdf_path = 'Bills/Invoice.pdf'
    data = extract_data_from_invoice(pdf_path)
    invoice_data = extract_invoice_data(data)
    print("Données de facture extraites:", invoice_data)

    if invoice_data:
        sheets_service = authenticate_google_sheets()
        sheet_id = '1A9UPxQ7uR6znZmZ96xJ7Klg2ycXKdFtGrSHmkVND2hs'
        update_google_sheet(sheets_service, sheet_id, invoice_data)
    else:
        print("Aucune donnée de facture extraite.")
