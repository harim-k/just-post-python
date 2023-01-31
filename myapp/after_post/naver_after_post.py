import myapp.utils.utils as utils
from myapp.after_post.after_post import *
from myapp.post.post import *


def save_excel_file(excel_file, 택배발송정보_dict):
    # get excel data
    excel_data = make_excel_data(excel_file, 택배발송정보_dict)

    # save as excel file
    save_as_excel_file(excel_data, '네이버_발송처리.xls')

    return excel_data


def make_excel_data(excel_file, 택배발송정보_dict):
    # load encrypted excel file
    decrypted_excel = load_encrypted_file(excel_file)

    # load decrypted excel file
    workbook = load_workbook(filename=decrypted_excel)

    # getting a customer sheet
    sheet = workbook.worksheets[0]

    # make after post excel data
    excel_data = list()

    # set first row from template
    excel_data.append(get_first_row_from_template())

    # set value row
    excel_data.extend(make_excel_value_data(sheet, 택배발송정보_dict))

    return excel_data


def make_excel_value_data(sheet, 택배발송정보_dict):
    excel_data = list()
    row_data = list()

    first_row = 3
    last_row = sheet.max_row
    # set value rows from store excel file
    for i in range(first_row, last_row + 1):
        row_data = make_excel_row_data(sheet, 택배발송정보_dict, str(i))

        if row_data == None:
            continue

        excel_data.append(row_data)

    return excel_data


def make_excel_row_data(sheet, 택배발송정보_dict, index):
    row_data = list()

    수취인명 = get_value_by_row_and_column(sheet, index, '수취인명')
    우편번호 = get_value_by_row_and_column(sheet, index, '우편번호')
    상품주문번호 = get_value_by_row_and_column(sheet, index, '상품주문번호')

    key = (수취인명, 우편번호)
    if key not in 택배발송정보_dict:
        return None

    배송방법 = "택배,등기,소포"
    택배사 = "CJ 대한통운"
    송장번호 = 택배발송정보_dict[key]

    row_data.append(상품주문번호)  # 상품주문번호
    row_data.append(배송방법)  # 배송방법
    row_data.append(택배사)  # 택배사
    row_data.append(송장번호)  # 송장번호

    return row_data


def get_first_row_from_template():
    row_data = list()
    workbook = load_workbook(filename="static/네이버_발송처리_템플릿.xlsx")

    # set first row from template
    for row in workbook.worksheets[0].values:
        for value in row:
            row_data.append(value)

    return row_data


def download_excel_file():
    return utils.download_excel_file('네이버_발송처리', 'xls')
