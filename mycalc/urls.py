from django.urls import path
from . import views
from django.conf.urls.static import static
from django.conf import settings

urlpatterns = [
    path('make_result_file/', views.make_result_file, name='make_result_file'),
    path('macro_run/', views.macro_run, name='macro_run'),
    path('add_trub/', views.add_trub, name='add_trub'),
    path('add_plosk/', views.add_plosk, name='add_plosk'),
    path('add_emk/', views.add_emk, name='add_emk'),

    path('main/', views.main, name='main'),
    path('form/', views.form, name='form'),
    path('form/js/0.js', views.index, name='index'),
    path('', views.index, name='index'),

    path('js/<str:page>', views.other_page_js, name='other_page_js'),
    path('form/js/<str:page>', views.other_page_form_js, name='other_page_form_js'),
    path('main/js/<str:page>', views.other_page_main_js, name='other_page_main_js'),
] + static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)