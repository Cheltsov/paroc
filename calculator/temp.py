from data import data_pipes as pipes
from data import data_planes as planes
from data import data_containers as containers


data_dicts = [pipes.data, planes.data, containers.data]


def input_in_dicts(data_dict):
    new_dict = {}
    for key in data_dict.keys():
        new_dict[key] = 'something'
    return new_dict


data_pipes = input_in_dicts(data_dicts[0])

print(data_pipes)
