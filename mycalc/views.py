from django.shortcuts import render, redirect
from django.http import HttpResponse, Http404
from django.template import TemplateDoesNotExist
from django.template.loader import get_template
from django.templatetags.static import static


# Create your views here.
def index(request):
    return render(request, 'mycalc/index.html')

def main(request):
    return render(request, 'mycalc/main.html')

def form(request):
    return render(request, 'mycalc/form.html')

def add(request):
    return ''

def other_page_js(request, page):
    return redirect('/static/mycalc/js/'+page)

def other_page_form_js(request, page):
    return redirect('/static/mycalc/js/'+page)

def other_page_main_js(request, page):
    return redirect('/static/mycalc/js/'+page)



