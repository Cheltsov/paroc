def input_in_dict(data_dict, data):
    new_dict = {}
    for key in data_dict.keys():
        new_dict[key] = data[key].value
    return new_dict


def input_in_sheet(sheet, data):
    for key, value in data.items():
        for row in range(len(data)):
            if sheet.cell(row=row + 2, column=1).value == key:
                sheet.cell(row=row + 2, column=2, value=value)
                break
