from util.excel_util import *
from util.io_util import *
import json


def build_json(title, data, name):
    json_str = '\t"' + name + '":{'
    if len(title) != len(data):
        print(data)
        print("本行数据和标头个数不一致")
    num = len(data)
    for i in range(0, num):
        if data[i] is not None:
            if name == 'value':
                item = '\n\t\t"' + title[i] + '":' + data[i] + ','
            else:
                item = '\n\t\t"' + title[i] + '":"' + data[i] + '",'
            json_str += item
    str_len = len(json_str)
    json_str = json_str[:str_len - 1]
    json_str += '\n\t},\n'
    return json_str


def build_script():
    script_docs = []
    condition_index = 6
    dimension_index = 13

    rows = read_excel_to_list('files\\test.xlsx', 1, 1, 6)
    title = rows[0]

    condition_title = title[:condition_index]
    dimension_title = title[condition_index:dimension_index]
    value_title = title[dimension_index:]

    data = rows[1:]
    for row in data:
        condition = row[:condition_index]
        dimension = row[condition_index:dimension_index]
        value = row[dimension_index:]

        condition_json = build_json(condition_title, condition, "condition")
        dimension_json = build_json(dimension_title, dimension, "dimension")
        value_json = build_json(value_title, value, "value")

        script = condition_json + dimension_json + value_json
        str_len = len(script)
        script = '{\n' + script[:str_len - 2] + "\n}\n\n"
        script_docs.append(script)
    # 所有json写入到json文件中，若需要每一行写入到单个文件  将写入操作移入到循环里
    write_docs_to_file(script_docs, 'files\\script.txt')


# build_script()

resp_json = json.load('files\\返回报文.json')
print(resp_json)