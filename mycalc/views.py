from django.shortcuts import render, redirect
from django.http import HttpResponse, Http404
from django.template import TemplateDoesNotExist
from django.template.loader import get_template
from django.templatetags.static import static
from paroc.settings import MEDIA_ROOT_W
import codecs


# Create your views here.
def index(request):
    return render(request, 'mycalc/index.html')

def main(request):
    f = codecs.open(MEDIA_ROOT_W + '\\regions.txt', "r", "utf_8_sig")
    regions = f.read().split('\r\n')
    regions = sorted(regions)

    f = codecs.open(MEDIA_ROOT_W + '\\insulations.txt', "r", "utf_8_sig")
    insulations = f.read().split('\r\n')

    f = codecs.open(MEDIA_ROOT_W + '\\insulations_plosk.txt', "r", "utf_8_sig")
    insulations_plosk = f.read().split('\r\n')

    context = {
        'regions': regions,
        'insulations': insulations,
        'insulations_plosk': insulations_plosk
    }
    return render(request, 'mycalc/main.html', context)

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



