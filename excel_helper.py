import xlrd


def open_excel_json(filename=None, sheet_index=0, start_index=0, select_column_index=None, select_column_name=None,
                    remove_enter=False):
    """
    :param filename: 文件名
    :param sheet_index: 读取的是第几个sheet
    :param start_index: 起始下标
    :param select_column_index: 根据第几列来获取数据，来默认全选,如需要获取指定列请传eg:[0,1]
    :param select_column_name: 根据表头来获取数据，默认全选,如需要获取指定列请传eg:["name","age"]
    :param remove_enter: 移除掉回车符号
    :return: 返回处理好的json list
    """
    workbook = xlrd.open_workbook(filename, encoding_override="utf-8")
    result_list = []
    sheet = workbook.sheet_by_index(sheet_index)
    row_len = sheet.nrows
    start = False
    header = []
    for row_index in range(row_len):
        data_array = sheet.row_values(row_index)
        if row_index == start_index:
            start = True
            header = data_array
        elif start is True:
            data = {}
            for column_index in range(len(data_array)):
                column_name = header[column_index]
                column_value = data_array[column_index]

                # 替换回车符号
                if remove_enter is True and type(column_value) is str:
                    column_value = column_value.replace("\n", "")

                if (select_column_index is None and select_column_name is None) or (
                        select_column_index is not None and column_index in select_column_index) or (
                        select_column_name is not None and column_name in select_column_name):
                    data[column_name] = column_value
            result_list.append(data)
    return result_list
