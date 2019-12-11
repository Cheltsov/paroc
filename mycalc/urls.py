from django.urls import path
from . import views
from django.conf.urls.static import static
from django.conf import settings

urlpatterns = [
    path('add/', views.add, name='add'),
    path('main/', views.main, name='main'),
    path('form/', views.form, name='form'),
    path('/form/js/0.js', views.index, name='index'),
    path('', views.index, name='index'),

    path('js/<str:page>', views.other_page_js, name='other_page_js'),
    path('form/js/<str:page>', views.other_page_form_js, name='other_page_form_js'),
    path('main/js/<str:page>', views.other_page_main_js, name='other_page_main_js'),
] + static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)