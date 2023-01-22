from django.urls import path

from . import views

app_name = "myapp"

urlpatterns = [
    path('', views.index, name='index'),

    path('make_post_excel_file',
         views.make_post_excel_file,
         name='make_post_excel_file'),

    path('make_after_post_excel_file',
         views.make_after_post_excel_file,
         name='make_after_post_excel_file'),

    path('download_gs_post_excel_file',
         views.download_gs_post_excel_file,
         name='download_gs_post_excel_file'),

    path('download_cu_post_excel_file',
         views.download_cu_post_excel_file,
         name='download_cu_post_excel_file'),

    path('download_naver_after_post_excel_file',
         views.download_naver_after_post_excel_file,
         name='download_naver_after_post_excel_file'),

    path('download_coupang_after_post_excel_file',
         views.download_coupang_after_post_excel_file,
         name='download_coupang_after_post_excel_file'),

    path('download_ably_after_post_excel_file',
         views.download_ably_after_post_excel_file,
         name='download_ably_after_post_excel_file'),

    path('merge_post_excel_file',
         views.merge_post_excel_file,
         name='merge_post_excel_file'),

]
