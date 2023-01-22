import myapp.post.cu_post as cu_post
import myapp.post.gs_post as gs_post
from myapp.utils.utils import *


def make_excel_file(request):
    excel_file = request.FILES["excel_file"]
    store_type = request.POST["store_type"]
    
    택배예약현황_string = request.POST["after_post_data"]
    택배발송정보_dict = get_택배발송정보_dict(택배예약현황_string)

    # save after post excel file
    excel_data = save_excel_file(excel_file, store_type, 택배발송정보_dict)

    return excel_data


def save_excel_file(excel_file, store_type, 택배발송정보_dict):
    import myapp.after_post.naver_after_post as naver_after_post
    import myapp.after_post.coupang_after_post as coupang_after_post
    import myapp.after_post.ably_after_post as ably_after_post

    if is_네이버(store_type):
        excel_data = naver_after_post.save_excel_file(excel_file, 택배발송정보_dict)
    elif is_쿠팡(store_type):
        excel_data = coupang_after_post.save_excel_file(excel_file, 택배발송정보_dict)
    elif is_에이블리(store_type):
        excel_data = ably_after_post.save_excel_file(excel_file, 택배발송정보_dict)

    return excel_data


def is_네이버(store_type):
    return store_type == '네이버'

def is_쿠팡(store_type):
    return store_type == '쿠팡'

def is_에이블리(store_type):
    return store_type == '에이블리'


def get_택배발송정보_dict(after_post_data_string):
    names, postcodes, invoice_numbers = get_택배발송정보(after_post_data_string)

    key = [(name, postcode) for name, postcode in zip(names, postcodes)]
    dict = {k: v for k, v in zip(key, invoice_numbers)}
    return dict


def get_택배발송정보(after_post_data_string):
    if is_cu_post(after_post_data_string):
        return cu_post.get_택배발송정보(after_post_data_string)
    else:
        return gs_post.get_택배발송정보(after_post_data_string)


def is_cu_post(after_post_data_string):
    return 'CUPOST' in after_post_data_string



def save_list_as_excel_file(excel_data):
    save_as_excel_file(excel_data, 'after_post.xlsx')
