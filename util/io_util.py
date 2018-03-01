"""
文件 IO 读写操作
"""


def write_docs_to_file(docs, dest_path, mode='w'):
    """
    将文档按照指定模式写入到文件中
    :param docs:
    :param dest_path:
    :param mode:
    :return:
    """
    try:
        with open(dest_path, mode=mode, encoding='UTF-8') as fis:
            for doc in docs:
                fis.write(doc + '\n')
    except Exception as e:
        print('写入文本文件失败')
        print(e)
    else:
        print('写入文本文件成功')

