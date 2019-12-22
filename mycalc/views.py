from django.shortcuts import render, redirect
from django.http import HttpResponse, Http404
from django.template import TemplateDoesNotExist
from django.template.loader import get_template
from django.templatetags.static import static
from paroc.settings import MEDIA_ROOT_W
import mycalc.additional_functions as af
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
    if request.method == 'POST':
        print(request.POST)
        return HttpResponse("post")
    else:
        return HttpResponse("no post")
    # from data import data_pipes as pipes
    # from data import data_planes as planes
    # from data import data_containers as containers
    # from openpyxl import load_workbook
    # import cgi
    # import cgitb
    #
    # cgitb.enable()  # for troubleshooting
    # data = cgi.FieldStorage()  # the cgi library gets vars from html
    #
    # sheet_names = ['Trub', 'Plosk', 'Emk']
    # data_dicts = [pipes.data, planes.data, containers.data]
    # num_of_sheets = len(sheet_names)
    #
    # empty_dicts = []
    # for i in range(num_of_sheets):
    #     empty_dicts.append(af.input_in_dict(data_dicts[i], data))
    #
    # filename = '../media/Калькулятор Парок ТИ 19_12_17.xlsm'  # todo заменить txt файл на xlsm
    # wb = load_workbook(filename=filename, read_only=False)
    #
    # for i in range(len(sheet_names)):
    #     sheet = wb.get_sheet_by_name(sheet_names[i])
    #     af.input_in_sheet(sheet, empty_dicts[i])
    #
    # wb.save(filename='../media/second-book.xlsx')


def other_page_js(request, page):
    return redirect('/static/mycalc/js/' + page)


def other_page_form_js(request, page):
    return redirect('/static/mycalc/js/' + page)


def other_page_main_js(request, page):
    return redirect('/static/mycalc/js/' + page)
