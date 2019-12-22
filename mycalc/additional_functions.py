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
