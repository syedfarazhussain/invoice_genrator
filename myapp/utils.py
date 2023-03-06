from django.contrib import messages
from .forms import UserDataForm
import openpyxl
import datetime
from docx2pdf import convert
from time import sleep


def process_de_data(request, lifting_file, user_data):
    wb = openpyxl.load_workbook(lifting_file, data_only=True)
    # get the first worksheet
    ws = wb.active
    lifting_file_index = ['Buy', 'Buy Amount', 'Sell', 'Sell Amount', 'Fee', 'Total Settlement', 'Beneficiary', 'When Booked', 'When Created', 'Created By', 'Delivery Method', 'Reference', 'Delivery Country ISO Code', 'Delivery Country Name', 'Your Reference', 'Cheque Number', 'Beneficiary ID', 'Execution Date', 'Payment Line Status', 'Payment Number', 'Processing Date', 'Bank Reference Number', 'AccountName', 'Bank Value Date', 'Exchange Rate', 'Our Reference', 'Charges Type', 'USD AMOUNT']

    if list(ws.columns) != lifting_file_index:
        messages.error(request, '!Invalid Lifting fee file.')
        return

    messages.info(request, '!Checking File')
    sleep(1)

    processed_data = {}
    for index, row in ws.iterrows():
        if pd.isna(row['AccountName']) == False:
            if row['AccountName'] in processed_data:
                processed_data[row['AccountName']].append(row.to_dict())
            else:
                processed_data[row['AccountName']] = [row.to_dict()]

    messages.info(request, '!Generating invoices')
    sleep(1)
    today = datetime.datetime.now()
    Word_To_PDF = convert(lifting_file)