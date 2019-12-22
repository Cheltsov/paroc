from django.shortcuts import render, redirect
from django.http import HttpResponse, Http404
from django.template import TemplateDoesNotExist
from django.template.loader import get_template
from django.templatetags.static import static
from paroc.settings import MEDIA_ROOT_W
import mycalc.additional_functions as af
from mycalc.data import data_pipes as pipes
from mycalc.data import data_planes as planes
from mycalc.data import data_containers as containers
from openpyxl import load_workbook
import codecs
import cgitb
import cgi


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
    if request.method == 'POST':
        dirty_data = request.POST
        data_type = request.POST.getlist('type')

        # Первый символ верхнего регистра в сооветствии с названиями таблиц
        sheet_name = data_type[0][0].upper() + data_type[0][1:]

        data = {}
        for k, v in dirty_data.items():
            data[k[5:-1]] = v
        print(data)

        data_dicts = {'Trub': pipes.data,
                      'Plosk': planes.data,
                      'Emk': containers.data
                      }

        empty_dict = af.input_in_dict(data_dicts[sheet_name], data)

        filename = 'media/Калькулятор Парок ТИ 19_12_17.xlsm'  # todo заменить txt файл на xlsm
        wb = load_workbook(filename=filename, read_only=False)

        sheet = wb.get_sheet_by_name(sheet_name)
        af.input_in_sheet(sheet, empty_dict)

        wb.save(filename='media/second-book.xlsx')

        return HttpResponse("post")
    else:
        return HttpResponse("no post")


def other_page_js(request, page):
    return redirect('/static/mycalc/js/' + page)


def other_page_form_js(request, page):
    return redirect('/static/mycalc/js/' + page)


def other_page_main_js(request, page):
    return redirect('/static/mycalc/js/' + page)
