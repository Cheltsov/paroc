import json
import os


def input_in_dict(data_dict, data):
    for key in data_dict.keys():
        if key in data:
            data_dict[key] = data[key]
        else:
            print("dict 'data' have not a key like " + key)
            data_dict[key] = None
    return data_dict


def input_in_sheet(sheet, data):
    for key, value in data.items():
        for row in range(len(data)):
            if sheet.cell(row=row + 2, column=1).value == key:
                sheet.cell(row=row + 2, column=2, value=value)
                break


def to_json(sheet, insulations, rows):
    ins = []
    for i in range(rows):
        by_id = {}
        _i = str(i + 2)
        tol = []
        for row in sheet['K' + _i:'V' + _i]:
            for cell in row:
                if cell.value is not None:
                    tol.append(str(cell.value))
        by_id['id'] = i
        by_id['name'] = insulations[i]
        by_id['tol'] = tol
        ins.append(by_id)

    return json.dumps(ins)


def remove_files(path_to_dir):
    filelist = [f for f in os.listdir(path_to_dir) if f.endswith(".xlsm")]
    count = 0
    for f in filelist:
        os.remove(os.path.join(path_to_dir, f))
        count += 1
    print('Удалено ' + str(count) + ' файлов с расштрением .xlsm')

    filelist = [f for f in os.listdir(path_to_dir) if f.endswith(".xlsx")]
    count = 0
    for f in filelist:
        os.remove(os.path.join(path_to_dir, f))
        count += 1
    print('Удалено ' + str(count) + ' файлов с расштрением .xlsx')
