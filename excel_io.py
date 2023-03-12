import os
import logging
import openpyxl
from collections import defaultdict
# from openpyxl.styles import Hyperlink

def get_history_aid(excel_path):
    aid_dict = defaultdict(list)
    if not os.path.isfile(excel_path):
        logging.warning(f'History excel not exist: {excel_path}')
        return aid_dict
    
    # 遍历 excel 获取每个account 的历史history
    # 打开Excel文件
    workbook = load_workbook(filename=excel_path)
    # 选择工作表
    sheet = workbook['Sheet1']
    # 获取表头
    header = [cell.value for cell in sheet[1]]
    # 获取公众号和文章ID列的索引
    gzh_index = header.index('公众号')
    article_id_index = header.index('文章ID')
    # 构建字典
    data_dict = {}
    for row in sheet.iter_rows(min_row=2, values_only=True):
        gzh_name = row[gzh_index]
        article_id = row[article_id_index]
        if gzh_name not in data_dict:
            data_dict[gzh_name] = [article_id]
        else:
            data_dict[gzh_name].append(article_id)
    return data_dict


def write_to_excel(file_path, data, sheet_name = 'Sheet1'):
    """将数据写入Excel文件"""
    # 检查文件是否存在，如果不存在则创建一个新的Excel文件
    if not os.path.isfile(file_path):
        workbook = openpyxl.Workbook()
        workbook.save(file_path)

    # 加载Excel文件
    workbook = openpyxl.load_workbook(file_path)

    # 检查sheet是否存在，如果不存在则创建一个新的sheet
    if sheet_name not in workbook.sheetnames:
        new_sheet = workbook.create_sheet(sheet_name)
    else:
        new_sheet = workbook[sheet_name]

    # 获取表头
    headers = []
    if new_sheet.max_row == 1:
        headers = list(data[0].keys())
        for col, header in enumerate(headers, start=1):
            new_sheet.cell(row=1, column=col, value=header)
        # 将列宽调整为自适应列宽
        for col in new_sheet.columns:
            max_length = 0
            column = col[0].column_letter  # 获取列字母
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2) * 1.2  # 为每个单元格添加留白，乘以1.2是为了更好的显示
            new_sheet.column_dimensions[column].width = adjusted_width
    else:
        for col in range(1, new_sheet.max_column + 1):
            headers.append(new_sheet.cell(row=1, column=col).value)

    num_rows_to_insert = len(data)
    num_existing_rows = new_sheet.max_row

    # 移动现有数据
    for row in range(num_existing_rows, 1, -1):
        for col in range(1, new_sheet.max_column + 1):
            cell = new_sheet.cell(row=row, column=col)
            new_row = row + num_rows_to_insert
            new_cell = new_sheet.cell(row=new_row, column=col)
            new_cell.value = cell.value

    # 插入新数据
    for row, item in enumerate(data, start=2):
        for col, value in item.items():
            col_index = headers.index(col) + 1
            if col == '文章链接':
                # 如果该列是“文章链接”，则将单元格设置为超链接
                hyperlink = openpyxl.utils.cell.coordinate_from_string(f'{openpyxl.utils.get_column_letter(col_index)}{row}')
                new_sheet[hyperlink].value = '链接'
                new_sheet[hyperlink].hyperlink = value
            else:
                new_sheet.cell(row=row, column=col_index, value=value)

    # 保存Excel文件
    workbook.save(file_path)



# def read_excel(file_path, sheet_name):
#     workbook = openpyxl.load_workbook(file_path)
#     sheet = workbook[sheet_name]
#     rows = list(sheet.rows)
#     headers = [cell.value for cell in rows[0]]

#     result = []
#     for row in rows[1:]:
#         item = {}
#         for col, cell in enumerate(row):
#             item[headers[col]] = cell.value
#         result.append(item)
#     return result

# def read_excel(file_path):
#     """_summary_

#     Args:
#         file_path (_type_): _description_

#     Returns:
#         Dict: [sheet_name]:[sheet_data]
#     """
#     workbook = openpyxl.load_workbook(file_path)
#     result = []
#     for sheet in workbook:
#         rows = list(sheet.rows)
#         headers = [cell.value for cell in rows[0]]
#         sheet_data = []
#         for row in rows[1:]:
#             item = {}
#             for col, cell in enumerate(row):
#                 item[headers[col]] = cell.value
#             sheet_data.append(item)
#         result.append({sheet.title: sheet_data})
#     return result


def read_all_columns(file_path, column_name):
    workbook = openpyxl.load_workbook(file_path)
    # 创建一个字典，用于存储每个sheet的数据
    data_dict = defaultdict(list)

    # 遍历所有sheet
    for sheet in workbook.worksheets:
        # 检查该sheet是否包含指定的列名
        col_index = 0
        for col in sheet.iter_cols(min_row=1, max_row=1, values_only=True):
            col_index += 1
            if column_name in col:
                break
        if sheet.cell(row=1, column=col_index).value != column_name:
            # 如果不包含，将该sheet的值设置为空列表
            data_dict[sheet.title] = []
            continue

        # 获取指定列的数据
        col_data = []
        for row in sheet.iter_rows(min_row=2, values_only=True):
            col_data.append(row[col_index - 1])

        # 将数据添加到字典中
        data_dict[sheet.title] = col_data

    # 返回包含所有sheet数据的字典
    return data_dict




if __name__ == '__main__':
    #search_aid('/Users/wangzepeng/Downloads/test.xlsx')
    # data = [
    # {"姓名": "A22A1", "年龄": 22, "性别": "男"},
    # {"姓名": "BB222", "年龄": 28, "性别": "男"},
    # {"姓名": "CC223", "年龄": 25, "性别": "女"}
    # ]

    # write_to_excel("example.xlsx", "Sheet", data)
    # write_to_excel("example.xlsx", "Sheet1", data)
    read_all_columns("./wechat_report.xlsx", 'aid')
    