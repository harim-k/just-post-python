from myapp.post.post import *
from myapp.utils.utils import *


def get_excel_data(excel_file):
    # load excel file
    workbook = load_workbook(filename=excel_file)

    # getting a customer sheet
    sheet = workbook.worksheets[1]

    return make_excel_data(sheet)


def make_excel_data(sheet):
    # make post excel data
    gs_excel_data = list()
    cu_excel_data = list()

    # set first row from template
    gs_excel_data.append(gs_post.get_first_row_from_template())
    cu_excel_data.append(cu_post.get_first_row_from_template())

    excel_data = make_excel_value_data(sheet)
    gs_excel_data.extend(excel_data)
    cu_excel_data.extend(excel_data)

    return gs_excel_data, cu_excel_data


def make_excel_value_data(sheet):
    excel_data = list()

    # set value rows from store excel file
    for i in range(2, sheet.max_row + 1):
        row_data = list()
        index_string = str(i)

        수취인명 = get_value_by_row_and_column(sheet, index_string, '수취인명')
        우편번호 = get_value_by_row_and_column(sheet, index_string, '우편번호')
        배송지주소 = get_value_by_row_and_column(sheet, index_string, '배송지 주소')

        수취인연락처 = get_value_by_row_and_column(sheet, index_string, '수취인 연락처')
        구매자연락처 = get_value_by_row_and_column(sheet, index_string, '연락처')

        수량 = get_value_by_row_and_column(sheet, index_string, '수량')
        배송메세지 = get_value_by_row_and_column(sheet, index_string, '배송 메모')

        옵션정보 = get_value_by_row_and_column(sheet, index_string, '옵션 정보')
        상품명 = get_value_by_row_and_column(sheet, index_string, '상품명')

        # 특수문자 제거
        옵션정보 = 옵션정보.replace('&', '')
        상품명 = 상품명.replace('&', '')

        # 옵션명 제거
        옵션정보 = 옵션정보.split('/')[0]

        품목 = 옵션정보 if 옵션정보 != ' ' else 상품명
        배송요청사항 = ' '.join([품목, 수량, 배송메세지])
        지불방법 = '선불'

        row_data.append(수취인명)  # 수취인명
        row_data.append(우편번호)  # 우편번호
        row_data.append(배송지주소)  # 주소 1
        row_data.append(배송지주소)  # 주소 2
        row_data.append(수취인연락처)  # 전화번호 (수취인)
        row_data.append(구매자연락처)  # 추가 전화번호
        row_data.append(배송요청사항)  # 배송요청사항
        row_data.append(지불방법)  # 지불방법

        # 같은 주소인 경우 하나로 합치기
        if has_same_address(excel_data, row_data):
            index = get_same_address_index(excel_data, row_data)
            # 배송요청사항에 품목, 수량 추가
            excel_data[index][6] = ' '.join(
                [품목, 수량, excel_data[index][6]])
        else:
            excel_data.append(row_data)

    return excel_data
