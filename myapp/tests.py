from pathlib import Path
from django.test import TestCase
from unittest import skip

import myapp.after_post.after_post as after_post
import myapp.post.post as post

import pandas as pd


class file():
    def __init__(self, name, content):
        self.name = name
        self.content = content


class YourTestClass(TestCase):

    # post test
    def test_naver_post(self):
        store_type = '네이버'

        self.post_test(store_type)

    def test_coupang_post(self):
        store_type = '쿠팡'

        self.post_test(store_type)

    def test_ably_post(self):
        store_type = '에이블리'

        self.post_test(store_type)

    def post_test(self, store_type):
        # given
        gs_excel_file = 'gs_post.xls'
        cu_excel_file = 'cu_post.xlsx'

        order_excel_file = f'test/{store_type}_orders.xlsx'
        output_gs_excel_file = f'test/{store_type}_{gs_excel_file}'
        output_cu_excel_file = f'test/{store_type}_{cu_excel_file}'

        # when
        post.save_excel_file(open(order_excel_file, 'rb'), store_type)

        # then
        self.assertTrue(compare_excel_files(
            gs_excel_file, output_gs_excel_file, 0))
        self.assertTrue(compare_excel_files(
            cu_excel_file, output_cu_excel_file, 0))



    # after post test

    def test_naver_after_post(self):
        store_type = '네이버'

        self.after_post_test(store_type)

    def test_coupang_after_post(self):
        store_type = '쿠팡'

        self.after_post_test(store_type)

    def test_ably_after_post(self):
        store_type = '에이블리'

        self.after_post_test(store_type)

    def after_post_test(self, store_type):
        
        # given
        extension = 'xls' if store_type == '네이버' else 'xlsx'
        sheet_index = 1 if store_type == '에이블리' else 0

        택배발송정보 = f'test/{store_type}_택배발송정보'
        order_excel_file = f'test/{store_type}_orders.xlsx'
        expected_excel_file = f'test/{store_type}_발송처리.{extension}'
        actual_excel_file = f'{store_type}_발송처리.{extension}'

        택배예약현황_string = Path(택배발송정보).read_text()
        택배발송정보_dict = after_post.get_택배발송정보_dict(택배예약현황_string)

        # when
        after_post.save_excel_file(open(order_excel_file, 'rb'),
                                   store_type, 택배발송정보_dict)

        # then
        self.assertTrue(compare_excel_files(
            actual_excel_file, expected_excel_file, sheet_index))


def compare_excel_files(file1, file2, sheet_index):
    df1 = pd.read_excel(file1, sheet_name=sheet_index)
    df2 = pd.read_excel(file2, sheet_name=sheet_index)
    return df1.equals(df2)
