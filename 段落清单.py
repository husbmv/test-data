import random
import openpyxl
from openpyxl.styles import Border, Side, Alignment, Font
from openpyxl.worksheet.pagebreak import Break
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.utils import get_column_letter
import pandas as pd
import math


def 设置每页一个表(sheet, 每个表行数, 每个表列数, 横向=False):
    """根据每个表所占行数，在行边界插入分页符，并设置 A4 100% 缩放。"""
    max_row = sheet.max_row
    sheet.row_breaks.brk = []
    if 每个表行数 > 0:
        表数量 = max_row // 每个表行数
        for i in range(1, 表数量):
            sheet.row_breaks.append(Break(id=i * 每个表行数))

    sheet.page_setup.paperSize = sheet.PAPERSIZE_A4
    sheet.page_setup.orientation = 'landscape' if 横向 else 'portrait'
    sheet.page_setup.scale = 100
    sheet.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=False)

    last_col_letter = get_column_letter(每个表列数)
    sheet.print_area = 'A1:{}{}'.format(last_col_letter, max_row)

# 定义一个函数，用于从完整字符串中提取前缀字符串
# 获取单项名称
def dxmc(full_str, prefix_str):
    if full_str.startswith(prefix_str):
        suffix = full_str[len(prefix_str):].strip("-")
        return suffix
    else:
        # 如果不匹配，尝试其他方法
        if "-" in full_str:
            return full_str.split("-")[-1]
        return ""


def one(sheet1, KS, total_rows, page_number, df,data_list):

    number_dd = data_list[1]
    number_ht = data_list[6]
    number_xm = data_list[4]
    name_xm = str(data_list[5])
    name_dd = str(data_list[2])
    name_sgdw = data_list[7]

    # 获取单项名称
    dxmc_2 = dxmc(name_dd, name_xm)



    # 光缆总长 = 2000
    # 地址 = "环城北路与云岫北路576光交至北路576光交华凌建材东墙"
    # 接头编号 = str(接头数) + "号"

    # 设置第一行行高
    sheet1.row_dimensions[KS + 1].height = 33

    # 设置 A2 至 A33 行高为 21.75
    for row_num in range(KS + 2, KS + 34):  # 从 KS+2（即A2的行号）遍历到31（即A31的行号）
        sheet1.row_dimensions[row_num].height = 20.5  # 设置行高

    

    # 假设以下变量已经定义并赋值：型号, 出厂盘号, 出厂长度, 光缆芯数, 光缆长度
    # 填写表头
    header_rows = [
        ["段落清单", "", "", "", "", ""],  # 第1行表头
        ["订单编号："+number_dd, "", "", "合同编号："+number_ht, "", ""],  # 第2行表头数据
        ["项目名称："+name_xm, "", "", "项目编号："+number_xm, "", ""],   # 第3行表头数据
        ["工程名称："+dxmc_2, "", "", "施工单位："+name_sgdw, "", ""],   # 第4行表头数据
        ["序号", "段落", "", "型号", "数量", "备注"],  # 第5行表头数据
    ]
    start_row = KS + 1
    for row_data in header_rows:
        for col_index, value in enumerate(row_data, start=1):
            sheet1.cell(row=start_row, column=col_index).value = value
        start_row += 1

    # 页尾
    footer_rows = [
        ["施工单位：签章", "", "监理单位：签章", "", "", ""],    # 第31行页尾
        ["日期：", "", "日期：", "", "", ""],                    # 第32行页尾
        ["建设单位：签章", "", "维护单位：签章", "", "", ""],     # 第33行页尾
        ["日期：", "", "日期：", "", "", ""]                    # 第34行页尾
    ]
    start_row1 = KS + 30
    for row_data in footer_rows:
        for col_index, value in enumerate(row_data, start=1):
            sheet1.cell(row=start_row1, column=col_index).value = value
        start_row1 += 1

       
    # 段落清单 
    total_rows = total_rows-24*page_number
 
    if total_rows > 24:
        total_data = 24
    else:
        total_data = total_rows
    
    for i in range(total_data):
        name_value = df.iloc[i+ page_number * 24,1]
        model_value = df.iloc[i+ page_number * 24,2]
        long_value = df.iloc[i+ page_number * 24,3]
        sheet1.cell(row=KS + i + 6, column=1).value = i + 1 + page_number * 24  # 序号
        sheet1.cell(row=KS + i + 6, column=2).value = name_value            # 段落名称
        sheet1.cell(row=KS + i + 6, column=4).value = model_value           # 型号
        sheet1.cell(row=KS + i + 6, column=5).value = int(long_value)/1000  # 光缆长度（公里）
        

    # 定义要合并的单元格区域
    merge_ranges = [
        "A{}:F{}".format(KS + 1, KS + 1),  # 合并第一行的A到J列
        "A{}:C{}".format(KS + 2, KS + 2),  # 合并第二行的A到J列
        "D{}:F{}".format(KS + 2, KS + 2),  # 合并第三行的A、B列
        "A{}:C{}".format(KS + 3, KS + 3),  # 合并第三行的C到E列
        "D{}:F{}".format(KS + 3, KS + 3),  # 合并第三行的F、G列
        "A{}:C{}".format(KS + 4, KS + 4),  # 合并第三行的H到J列
        "D{}:F{}".format(KS + 4, KS + 4),  # 合并第四行到第五行的A列
        "A{}:B{}".format(KS + 30, KS + 30),  # 合并第四行到第五行的F列
        "C{}:F{}".format(KS + 30, KS + 30),  # 合并第四行的B到E列
        "A{}:B{}".format(KS + 31, KS + 31),  # 合并第四行到第五行的F列
        "C{}:F{}".format(KS + 31, KS + 31),  # 合并第四行的B到E列
        "A{}:B{}".format(KS + 32, KS + 32),  # 合并第四行到第五行的F列
        "C{}:F{}".format(KS + 32, KS + 32),  # 合并第四行的B到E列
        "A{}:B{}".format(KS + 33, KS + 33),  # 合并第四行到第五行的F列
        "C{}:F{}".format(KS + 33, KS + 33),  # 合并第四行的B到E列
    ]
    
    # 循环遍历合并单元格
    for merge_range in merge_ranges:
        sheet1.merge_cells(merge_range)  # 调用合并单元格的方法，将指定范围的单元格合并
     # 循环遍历合并单元格   
    for row in range(KS + 5, KS + 30):
        sheet1.merge_cells(start_row=row, start_column=2, end_row=row, end_column=3)

    # 设置A1字体、大小、加粗、自动换行、水平居中
    font = Font(name='宋体', size=20, bold=True)
    sheet1.cell(row=KS + 1, column=1).font = font
    sheet1.cell(row=KS + 1, column=1).alignment = Alignment(wrap_text=True,horizontal= 'center', vertical='center')
    


    # 设置 A2:F31 区域的文字上下左右居中
    alignment_center = Alignment(horizontal='center', vertical='center')        # 设置文字居中
    vertical_center = Alignment(vertical='center')                              # 设置文字垂直居中
    left_alignment = Alignment(horizontal='left', vertical='center')            # 设置文字左对齐

    for col in range(1, 7):
        for row in range(KS + 5, KS + 30):
            cell = sheet1.cell(row=row, column=col)
            cell.alignment = alignment_center           # 居中
            cell.font = Font(name='宋体', size=8)
        for row in range(KS + 2, KS + 5):
            cell = sheet1.cell(row=row, column=col)
            cell.alignment = vertical_center            # 设置单元格格式，垂直居中
            cell.font = Font(name='宋体', size=8)
        for row in range(KS + 30, KS + 34):
            cell = sheet1.cell(row=row, column=col)
            cell.alignment = vertical_center            # 设置单元格格式，垂直居中
            cell.font = Font(name='宋体', size=8)
        
    for row in range(KS + 6, KS + 30):
        cell = sheet1.cell(row=row, column=2)
        cell.alignment = left_alignment               # 左对齐     
    # 设置单元格格式，保留3位小数
    number_format = ('0.000')

    # B7:B30 和 E7:E30 范围内的单元格设置格式
    for row in range(KS + 6, KS + 30):  
        cell = sheet1.cell(row=row, column=5)
        cell.number_format = number_format
    

    # 添加边框线
    border = Border(left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin'))
    for row in range(KS + 5, KS + 30):
        for col in range(1, 7):
            cell = sheet1.cell(row=row, column=col)
            cell.border = border


# # 开始执行主程序
def zxcx_one(filter,data_list_one):
    # 创建一个新的工作簿
    wb = openpyxl.Workbook()

    # 获取默认的活动工作表
    sheet1 = wb.active

    KS = 0
    data_list = data_list_one


    # 设置 A1 至 J1 列的宽度为 6.75
    # 设置每列的指定宽度
    column_widths = [4.28, 37.65, 13.15, 12.03, 8.9, 9.03]
    for i, width in enumerate(column_widths, start=1):
        column_letter = openpyxl.utils.get_column_letter(i)
        sheet1.column_dimensions[column_letter].width = width
        
    # round_name='批量打印模板（不用打印1）.xlsx'

    df = pd.read_excel(filter, sheet_name='数据')

    # df = df.iloc[:,[1,2,3,4]]
    # 计算df列表的总行数
    total_rows = df.shape[0]

    mc = df.iloc[0,1]
    page_number = math.ceil(total_rows / 24)    # 计算页数

    # print(mc)
    # print(total_rows)
    # print(df)
    for i in range(page_number):


        row = sheet1.max_row
        if row == 1:
            max_row = 0
        else:
            max_row = row

        one(sheet1, max_row, total_rows, i, df,data_list)

    # 每张表占 34 行、6 列,A4 纵向,每张表独占一页
    设置每页一个表(sheet1, 每个表行数=34, 每个表列数=6, 横向=False)

    # 保存工作簿
    名称 = "段落清单.xlsx"
    wb.save(名称)