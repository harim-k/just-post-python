from openpyxl import load_workbook

import myapp.utils.utils as utils


def get_택배발송정보(post_data_string):

    strings = post_data_string.replace('\r\n', '')
    strings = strings.replace('\n', '')
    strings = strings.split('이름')[2:]

    names = [string.split('전화번호')[0] for string in strings]
    postcodes = [string.split('[')[1].split(']')[0] for string in strings]
    invoice_numbers = [string.split('운송장번호')[1].split('배송조회')[0] for string in strings]

    return names, postcodes, invoice_numbers


def get_first_row_from_template():
    workbook = load_workbook(filename='static/cu_post_template.xlsx')
    row_data = list()

    for row in workbook.worksheets[0].values:
        for value in row:
            row_data.append(value)

    return row_data


def save_as_excel_file(post_excel_data):
    utils.save_as_excel_file(post_excel_data, 'cu_post.xlsx')


def download_post_excel_file(request):
    return utils.download_excel_file('cu_post', 'xlsx')
