from openpyxl import load_workbook
from docx import Document
import datetime

def main():
    companies_data = {}
    template = 'temp.docx'
    result = 'result.docx'
    wb = load_workbook('data_isukk_8.xlsx', data_only=True)
    ws = wb['27.03.2022']
    colB = ws['B']
    for val in colB:
        if val.value == 'Грузополучатель':
            continue
        if val.value not in companies_data:
            companies_data[val.value] = {'Contracts': {'Contract1': {}, 'Contract2': {}, 'Contract3': {}},
                                         'Information': {'{{SENDER}}': None, '{{RECEIVER}}': None, '{{REASON}}' : None, '{{DATE}}' :  None}}

    for row in ws.iter_rows(min_row=3, values_only=True):
        sender = row[0]
        company = row[1]
        reason = row[2]
        product = row[3]
        unit = row[4]
        price = row[5]
        count = row[6]
        summa = row[7]
        if (len(companies_data[company]['Contracts']['Contract1']) == 0):
            companies_data[company]['Information']['{{SENDER}}'] = sender
            companies_data[company]['Information']['{{RECEIVER}}'] = company
            companies_data[company]['Contracts']['Contract1']['{{PRODUCT}}'] = product
            companies_data[company]['Contracts']['Contract1']['{{UNIT}}'] = unit
            companies_data[company]['Contracts']['Contract1']['{{PRICE}}'] = price
            companies_data[company]['Contracts']['Contract1']['{{COUNT}}'] = count
            companies_data[company]['Contracts']['Contract1']['{{SUMM}}'] = summa
            companies_data[company]['Information']['{{REASON}}'] = reason
            continue

        elif (len(companies_data[company]['Contracts']['Contract2']) == 0):
            companies_data[company]['Information']['SENDER'] = sender
            companies_data[company]['Information']['RECEIVER'] = company
            companies_data[company]['Contracts']['Contract2']['PRODUCT'] = product
            companies_data[company]['Contracts']['Contract2']['UNIT'] = unit
            companies_data[company]['Contracts']['Contract2']['PRICE'] = price
            companies_data[company]['Contracts']['Contract2']['COUNT'] = count
            companies_data[company]['Contracts']['Contract2']['SUMM'] = summa
            companies_data[company]['Information']['REASON'] = reason

        else:
            companies_data[company]['Information']['SENDER'] = sender
            companies_data[company]['RECEIVER'] = company
            companies_data[company]['Contracts']['Contract3']['PRODUCT'] = product
            companies_data[company]['Contracts']['Contract3']['UNIT'] = unit
            companies_data[company]['Contracts']['Contract3']['PRICE'] = price
            companies_data[company]['Contracts']['Contract3']['COUNT'] = count
            companies_data[company]['Contracts']['Contract3']['SUMM'] = summa
            companies_data[company]['Information']['REASON'] = reason

    template_doc = Document(template)
    for comp in companies_data:
        for key, value in companies_data[comp]['Information'].items():
            for paragraph in template_doc.paragraphs:
                replace_text(paragraph, key, value)

    template_doc.save(result)


def replace_text(paragraph, key, value):
    if key in paragraph.text:
        paragraph.text = paragraph.text.replace(key, value)


main()