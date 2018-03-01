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


def build_response(file):
    docs = []
    # response.xlsx 第一个sheet的第一行存返回的字段名
    titles = read_excel_to_list('files\\response.xlsx', 1, 1, 1)[0]
    docs.append(titles)
    with open(file, 'r') as f:
        resp = json.load(f)
    for doc in resp:
        condition_dict = doc["condition"]
        dimension_dict = doc["dimension"]
        value_dict = doc["value"]
        resp_dict = dict(condition_dict, **dimension_dict, **value_dict)
        item = []
        for title in titles:
            if resp_dict.__contains__(title):
                item.append(resp_dict[title])
            else:
                item.append("")
        docs.append(item)
    write_list_to_excel("files\\response_xlsx.xlsx", docs, "返回excel")


# build_script()
build_response("files\\返回报文.json")
