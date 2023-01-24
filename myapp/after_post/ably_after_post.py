import myapp.utils.utils as utils
from myapp.after_post.after_post import *
from myapp.post.post import *


def save_excel_file(excel_file, 택배발송정보_dict):
    # load excel file
    workbook = load_workbook(filename=excel_file)

    # getting a customer sheet
    sheet = workbook.worksheets[1]

    excel_data = make_excel_data(sheet, 택배발송정보_dict)

    workbook.save('에이블리_발송처리.xlsx')

    return excel_data


def make_excel_data(sheet, 택배발송정보_dict):
    # make after post excel data
    excel_data = list()

    # set first row from template
    excel_data.append(get_first_row_from_template())

    # set value row
    excel_data.extend(make_excel_value_data(sheet, 택배발송정보_dict))

    return excel_data


def make_excel_value_data(sheet, 택배발송정보_dict):
    excel_data = list()

    first_row = 2
    last_row = sheet.max_row
    # set value rows from store excel file
    for i in range(first_row, last_row + 1):
        row_data = make_excel_row_data(sheet, 택배발송정보_dict, str(i))

        if row_data == None:
            continue

        excel_data.append(row_data)

    return excel_data


def make_excel_row_data(sheet, 택배발송정보_dict, row_index):
    row_data = [cell.value for cell in sheet[row_index]]
    송장번호_index = get_index_by_row_and_column(sheet, row_index, '송장번호')

    수취인명 = get_value_by_row_and_column(sheet, row_index, '수취인명')
    우편번호 = get_value_by_row_and_column(sheet, row_index, '우편번호')

    key = (수취인명, 우편번호)

    if key not in 택배발송정보_dict:
        return None

    송장번호 = 택배발송정보_dict[key]

    # 5번째 값이 송장번호
    row_data[5] = 송장번호
    sheet[송장번호_index] = 송장번호

    return row_data


def get_first_row_from_template():
    row_data = list()
    workbook = load_workbook(
        filename="static/ably_after_post_template.xlsx")

    # set first row from template
    for row in workbook.worksheets[1].values:
        for value in row:
            row_data.append(value)

    return row_data


def download_excel_file():
    return utils.download_excel_file('에이블리_발송처리', 'xlsx')
