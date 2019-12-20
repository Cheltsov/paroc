from data import data_pipes as pipes
from data import data_planes as planes
from data import data_containers as containers

import cgi, cgitb

cgitb.enable()  # for troubleshooting
data = cgi.FieldStorage()  # the cgi library gets vars from html

sheet_names = ['Trub', 'Plosk', 'Emk']
data_dicts = [pipes.data, planes.data, containers.data]


def input_in_dict(data_dict):
    new_dict = {}
    for key in data_dict.keys():
        new_dict[key] = data[key].value
    return new_dict


data_pipes = input_in_dict(data_dicts[0])
data_planes = input_in_dict(data_dicts[1])
data_containers = input_in_dict(data_dicts[2])

