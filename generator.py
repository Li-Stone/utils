from util.excel_util import *
from util.io_util import *
import json
import requests


def build_json(title, data, name):
    """
    根据请求参数值列表和对应的参数名列表 拼接为请求的json字符串
    :param title: 请求参数名列表
    :param data:  请求参数值列表
    :param name:  请求参数所属分类（condition|dimension|value）
    :return:
    """
    json_str = '\t\t"' + name + '":{'
    if len(title) != len(data):
        print(data)
        print("本行数据和标头个数不一致")
    num = len(data)
    for i in range(0, num):
        # 忽略值为空的参数
        if data[i] is not None and str(data[i]).strip("") != '':
            if name == 'value':
                item = '\n\t\t\t"' + title[i] + '":' + data[i] + ','
            else:
                # 从excel中读取的日期格式数据 默认在后面带有时分秒 去掉  否则请求不成功
                item = '\n\t\t\t"' + title[i] + '":"' + str(data[i]).rstrip(" 00:00:00") + '",'
            json_str += item
    str_len = len(json_str)
    json_str = json_str[:str_len - 1]
    json_str += '\n\t\t},\n'
    return json_str


def build_script(req_excel, req_txt, resp_json):
    """
    将请求数据excel进行解析，循环请求，并将请求报文写到req_txt 指定的文件中（默认为覆盖）
    :param resp_json:   返回报文的保存路径
    :param req_excel:  请求数据excel路径
    :param req_txt:   生成的请求报文txt路径
    :return:
    """
    script_docs = ["["]
    response_xml = ["["]
    # condition 的最后一个请求参数在excel中列的位置
    condition_index = 6
    # dimension 的最后一个请求参数在excel中列的位置
    dimension_index = 36
    # value 的最后一个请求参数在excel中列的位置
    value_index = 43

    rows = read_excel_to_list(req_excel, 1, 1, 52)
    title = rows[0]

    # 将请求参数的参数名 按照condition|dimension|value  分成三段
    condition_title = title[:condition_index]
    dimension_title = title[condition_index:dimension_index]
    value_title = title[dimension_index:value_index]

    data = rows[1:]
    results = []
    for row in data:
        # 从请求数据列表中截取 condition|dimension|value 对应的值
        condition = row[:condition_index]
        dimension = row[condition_index:dimension_index]
        value = row[dimension_index:value_index]

        # 根据各自的请求参数值列表  和  请求参数名列表 生成请求json串
        condition_json = build_json(condition_title, condition, "condition")
        dimension_json = build_json(dimension_title, dimension, "dimension")
        value_json = build_json(value_title, value, "value")

        # 将请求json串拼接为最终的请求串
        script = condition_json + dimension_json + value_json
        str_len = len(script)
        script = '{\n\t"auth": {},\n\t "channel": {},\n\t "log": {},\n\t ' \
                 '"method": "queryLifeTableData",\n\t "param":{\n' + \
                 script[:str_len - 2] + '\n\t},\n\t"version":""\n}'

        # script 为字符串，需要转换为dict类型 进行请求
        param = json.loads(script)

        r = requests.post('http://10.8.198.107:8080/iltplatformMpp/lifetablecompilefacade', json=param)
        response_xml.append("," + r.text)
        result = json.loads(r.text)

        results.append(result["data"]["resultData"][0])
        script_docs.append("," + script + ",")
    # 将请求和响应报文 调整为标准对象数组格式：  删除第一个报文前的逗号
    response_xml.append("]")
    script_docs.append("]")
    script_docs[1] = script_docs[1][1:]
    response_xml[1] = response_xml[1][1:]
    # 所有json写入到json文件中，若需要每一行写入到单个文件  将写入操作移入到循环里
    write_docs_to_file(script_docs, req_txt)
    write_docs_to_file(response_xml, resp_json)
    return results


def build_response(results, resp_title, resp_excel):
    """
    根据返回结果的报文列表，解析数据写入到excel
    :param results:  所有请求的返回结果报文列表
    :param resp_title:  返回结果报文字段的excel路径
    :param resp_excel:  返回结果Excel的保存路径
    :return:
    """
    docs = []
    # resp_title 第一个sheet的第一行存返回的字段名
    titles = read_excel_to_list(resp_title, 1, 1, 1)[0]
    docs.append(titles)
    for doc in results:
        item = []
        for title in titles:
            if doc.__contains__(title):
                item.append(doc[title])
            else:
                item.append("")
        docs.append(item)
    write_list_to_excel(resp_excel, docs, "测试返回数据")


request_excel = "files\\请求数据.xlsx"  # 请求数据路径
request_json = "files\\请求报文.json"  # 请求报文路径
response_title = "files\\返回字段.xlsx"  # 返回字段excel路径
response_excel = "files\\返回数据.xlsx"  # 返回数据excel路径
response_json = "files\\返回报文.json"  # 返回报文json路径


result_json = build_script(request_excel, request_json, response_json)  # 获取响应报文中的目标数据
build_response(result_json, response_title, response_excel)  # 解析响应报文中的目标数据并将解析后的数据写入到excel中
