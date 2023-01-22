from django.shortcuts import render

import myapp.after_post.ably_after_post as ably_after_post
import myapp.after_post.after_post as after_post
import myapp.after_post.coupang_after_post as coupang_after_post
import myapp.after_post.naver_after_post as naver_after_post
import myapp.post.cu_post as cu_post
import myapp.post.gs_post as gs_post
import myapp.post.post as post
import myapp.utils.utils as utils


def index(request):
    return render(request, 'myapp/index.html', {})


def make_post_excel_file(request):
    if "POST" != request.method:
        return render(request, 'myapp/index.html', {})

    return render(request, 'myapp/index.html',
                  {"post_excel_data": post.make_excel_file(request)})


def make_after_post_excel_file(request):
    if "POST" != request.method:
        return render(request, 'myapp/index.html', {})

    return render(request, 'myapp/index.html',
                  {"after_post_excel_data": after_post.make_excel_file(request)})


def download_gs_post_excel_file(request):
    return gs_post.download_post_excel_file(request)


def download_cu_post_excel_file(request):
    return cu_post.download_post_excel_file(request)


def download_naver_after_post_excel_file(request):
    return naver_after_post.download_excel_file()


def download_coupang_after_post_excel_file(request):
    return coupang_after_post.download_excel_file()


def download_ably_after_post_excel_file(request):
    return ably_after_post.download_excel_file()


def merge_post_excel_file(request):
    return utils.merge_post_excel_file(request)
