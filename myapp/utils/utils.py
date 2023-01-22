import datetime
import pytz

import io
import msoffcrypto
from openpyxl import load_workbook

import pandas as pd

from django.http import HttpResponse


KST = pytz.timezone('Asia/Seoul')
PASSWORD = '1111'


def get_today():
    return str(datetime.datetime.now(KST)).split(' ')[0]


def download_after_post_excel_file(new_filename, extension):
    with open(f'after_post.xlsx', 'rb') as f:
        filename = f'{new_filename}_{get_today()}.{extension}'

        response = HttpResponse(f, content_type='application/ms-excel')
        response['Content-Disposition'] = f'attachment; filename={filename}'

    return response


def download_excel_file(origin_filename, extension):
    with open(f'{origin_filename}.{extension}', 'rb') as f:
        filename = f'{origin_filename}_{get_today()}.{extension}'

        response = HttpResponse(f, content_type='application/ms-excel')
        response['Content-Disposition'] = f'attachment; filename={filename}'

    return response


def save_as_excel_file(excel_data, filename):
    pd.DataFrame(excel_data).to_excel(
        filename, sheet_name='발송처리', index=False, header=False)


def load_encrypted_file(excel_file):
    decrypted_excel = io.BytesIO()

    office_file = msoffcrypto.OfficeFile(excel_file)
    office_file.load_key(password=PASSWORD)
    office_file.decrypt(decrypted_excel)

    return decrypted_excel


def merge_post_excel_file(request):
    excel_files = request.FILES.getlist("excel_file")

    sheets = [pd.read_excel(excel_file, sheet_name=0, dtype=pd.StringDtype())
              for excel_file in excel_files]

    merged_sheet = pd.concat(sheets)

    merged_sheet.to_excel('merged_excel_file.xls',
                          sheet_name='발송처리', index=False, header=True)

    return download_excel_file('merged_excel_file', 'xls')
