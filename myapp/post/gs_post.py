from openpyxl import load_workbook

import myapp.utils.utils as utils


def get_택배발송정보(post_data_string):

    strings = post_data_string.replace('\r\n', '')
    strings = strings.replace('\n', '')
    strings = strings.split('수신정보')[1]
    strings = strings.split('선불')[:-1]

    delimeter = '반품' if '반품' in strings[0] else 'Address'  # 사업자 계정 유무에 따라 다름

    names = [string.split(delimeter)[0] for string in strings]
    postcodes = [string.split('[')[1].split(']')[0] for string in strings]
    invoice_numbers = [string.split('운송장번호')[1].split('Comment')[0] for string in strings]

    return names, postcodes, invoice_numbers


def get_first_row_from_template():
    workbook = load_workbook(filename='static/gs_대량발송_템플릿.xlsx')
    row_data = list()

    for row in workbook.worksheets[0].values:
        for value in row:
            row_data.append(value)

    return row_data


def save_as_excel_file(post_excel_data):
    utils.save_as_excel_file(post_excel_data, 'gs_대량발송.xls')


def download_post_excel_file(request):
    return utils.download_excel_file('gs_대량발송', 'xls')
