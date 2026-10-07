import random
from openpyxl.styles import Border, Side, Alignment, Font
from openpyxl.worksheet.pagebreak import Break
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.utils import get_column_letter
import pandas as pd


def 设置每页一个表(sheet, 每个表行数, 每个表列数, 横向=True):
    """根据每个表所占行数，在行边界插入分页符，并设置 A4 适应一页。"""
    max_row = sheet.max_row
    sheet.row_breaks.brk = []
    if 每个表行数 > 0:
        表数量 = max_row // 每个表行数
        for i in range(1, 表数量):
            sheet.row_breaks.append(Break(id=i * 每个表行数))

    sheet.page_setup.paperSize = sheet.PAPERSIZE_A4
    sheet.page_setup.orientation = 'landscape' if 横向 else 'portrait'
    sheet.page_setup.scale = 100  # 100% 缩放，不压缩；分页符保证每张表独占一页
    sheet.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=False)

    last_col_letter = get_column_letter(每个表列数)
    sheet.print_area = 'A1:{}{}'.format(last_col_letter, max_row)


def 第一个表(sheet1, 光缆芯数, KS, 光缆总长, 地址):

    # 光缆总长 = 2000
    # 地址 = "环城北路与云岫北路576光交至北路576光交华凌建材东墙"
    

    # 设置第一行行高
    sheet1.row_dimensions[KS + 1].height = 45.75
    sheet1.row_dimensions[KS + 2].height = 21.75
    sheet1.row_dimensions[KS + 3].height = 21.75
    sheet1.row_dimensions[KS + 4].height = 21.75
    sheet1.row_dimensions[KS + 21].height = 21.75

    # 设置 A2 至 A33 行高为 21.75
    for row_num in range(KS + 5, KS + 21):  # 从 KS+2（即A2的行号）遍历到31（即A31的行号）
        sheet1.row_dimensions[row_num].height = 33.75  # 设置行高

    

    # 假设以下变量已经定义并赋值：型号, 出厂盘号, 出厂长度, 光缆芯数, 光缆长度
    # 填写表头
    header_rows = [
        ["表四 "+ 地址 +"用户段光纤线路衰减测试记录", "", "", "", "", "", "", "", "", "", "", ""],  # 第2行表头
        ["OTDR:s20C/N          折射率:1.4659          波长:1550 nm          光缆长度:" + 光缆总长 + "km", "", "", "", "", "", "", "", "", "", "", ""],  # 第3行表头数据
        ["光纤序号", "", "损耗","", "光纤序号", "", "损耗", "","光纤序号", "", "损耗", ""],  # 第4行表头数据
        ["", "", "db", "db/km", "", "", "db", "db/km","", "", "db", "db/km"],  # 第5行表头数据
    ]
    start_row = KS + 1
    for row_data in header_rows:
        for col_index, value in enumerate(row_data, start=1):
            sheet1.cell(row=start_row, column=col_index).value = value
        start_row += 1

    # 页尾
    footer_rows = [
        ["折射波长：1550nm   测试人：                  监理：                日期：          ", "", "", "", "", "", "", "", "", ""]  # 第33行页尾
    ]
    start_row1 = KS + 21
    for row_data in footer_rows:
        for col_index, value in enumerate(row_data, start=1):
            sheet1.cell(row=start_row1, column=col_index).value = value
        start_row1 += 1

    def 损耗值():
        return round(random.uniform(0.141, 0.131), 3)
    
    def 平均损耗值():
        return round(random.uniform(0.202, 0.191), 3)
    
    for m in range(16):
        # 设置序号
        sheet1.cell(row=KS + m + 5, column=1).value = m + 1
        sheet1.cell(row=KS + m + 5, column=5).value = m + 17
        sheet1.cell(row=KS + m + 5, column=9).value = m + 33
        # 设置A-B
        sheet1.cell(row=KS + m + 5, column=2).value = "A-B"
        sheet1.cell(row=KS + m + 5, column=6).value = "A-B"
        sheet1.cell(row=KS + m + 5, column=10).value = "A-B"
    
    
    if 光缆芯数 in [48, 96, 144]:
        for m in range(16):
            loss_value = 损耗值()
            loss_value_1 = 平均损耗值()
            # 设置损耗值
            sheet1.cell(row=KS + m + 5, column=3).value = loss_value
            sheet1.cell(row=KS + m + 5, column=7).value = loss_value
            sheet1.cell(row=KS + m + 5, column=11).value = loss_value
            # 设置平均损耗值
            sheet1.cell(row=KS + m + 5, column=4).value = loss_value_1
            sheet1.cell(row=KS + m + 5, column=8).value = loss_value_1
            sheet1.cell(row=KS + m + 5, column=12).value = loss_value_1

    elif 光缆芯数 == 24:
        for m in range(8):
            loss_value = 损耗值()
            loss_value_1 = 平均损耗值()
            # 设置损耗值
            sheet1.cell(row=KS + m + 5, column=3).value = loss_value
            sheet1.cell(row=KS + m + 13, column=3).value = loss_value
            sheet1.cell(row=KS + m + 5, column=7).value = loss_value
            # 设置平均损耗值
            sheet1.cell(row=KS + m + 5, column=4).value = loss_value_1
            sheet1.cell(row=KS + m + 13, column=4).value = loss_value_1
            sheet1.cell(row=KS + m + 5, column=8).value = loss_value_1

    else:
        for i in [4, 8, 12, 24]:
            if 光缆芯数 == i:
                for m in range(i):
                    loss_value = 损耗值()
                    loss_value_1 = 平均损耗值()
                    # 设置损耗值
                    sheet1.cell(row=KS + m + 5, column=3).value = loss_value                     
                    # 设置平均损耗值
                    sheet1.cell(row=KS + m + 5, column=4).value = loss_value_1

                break
        # else:
        # print("Invalid input")

    # 定义要合并的单元格区域
    merge_ranges = [
        "A{}:L{}".format(KS + 1, KS + 1),  # 合并第一行的A到J列
        "A{}:L{}".format(KS + 2, KS + 2),  # 合并第二行的A到J列
        "A{}:B{}".format(KS + 3, KS + 4),  # 合并第三行的A、B列
        "E{}:F{}".format(KS + 3, KS + 4),  # 合并第三行的A、B列
        "I{}:J{}".format(KS + 3, KS + 4),  # 合并第三行的A、B列
        "C{}:D{}".format(KS + 3, KS + 3),  # 合并第三行的A、B列
        "G{}:H{}".format(KS + 3, KS + 3),  # 合并第三行的A、B列
        "K{}:L{}".format(KS + 3, KS + 3),  # 合并第三行的A、B列
        "A{}:L{}".format(KS + 21, KS + 21)  # 合并第21行的A到J列
    ]
    
    # 循环遍历合并单元格
    for merge_range in merge_ranges:
        sheet1.merge_cells(merge_range)  # 调用合并单元格的方法，将指定范围的单元格合并

    # 设置A1字体、大小、加粗、自动换行、水平居中
    font = Font(name='宋体', size=14, bold=True)
    sheet1.cell(row=KS + 1, column=1).font = font
    sheet1.cell(row=KS + 1, column=1).alignment = Alignment(wrap_text=True,horizontal= 'center', vertical='center')
    


        # 设置 A2:F31 区域的文字上下左右居中
    alignment_center = Alignment(horizontal='center', vertical='center')
    for row in range(KS + 2, KS + 22):
        for col in range(1, 13):
            cell = sheet1.cell(row=row, column=col)
            cell.alignment = alignment_center

            # 设置单元格格式，保留3位小数
    number_format = ('0.000')

    # B7:B30 和 E7:E30 范围内的单元格设置格式
    for row in range(KS + 5, KS + 21):
        for col in [3, 4]:  # 列C和列D的列号
            cell = sheet1.cell(row=row, column=col)
            cell.number_format = number_format
        for col in [7, 8]:  # 列G和列H的列号
            cell = sheet1.cell(row=row, column=col)
            cell.number_format = number_format
        for col in [11, 12]:  # 列K和列L的列号
            cell = sheet1.cell(row=row, column=col)
            cell.number_format = number_format

    

    # 添加边框线
    border = Border(left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin'))
    for row in range(KS + 3, KS + 21):
        for col in range(1, 13):
            cell = sheet1.cell(row=row, column=col)
            cell.border = border

def 第二个表(sheet1, 光缆芯数, KS, 光缆总长, 地址):

    # 光缆总长 = 2000
    # 地址 = "环城北路与云岫北路576光交至北路576光交华凌建材东墙"

    bgs = 1 * 21  # 前面表格数量

    # 设置第一行行高
    sheet1.row_dimensions[bgs + KS + 1].height = 45.75
    sheet1.row_dimensions[bgs + KS + 2].height = 21.75
    sheet1.row_dimensions[bgs + KS + 3].height = 21.75
    sheet1.row_dimensions[bgs + KS + 4].height = 21.75
    sheet1.row_dimensions[bgs + KS + 21].height = 21.75

    # 设置 A2 至 A33 行高为 21.75
    for row_num in range(bgs + KS + 5, bgs + KS + 21):  # 从 bgs + KS+2（即A2的行号）遍历到31（即A31的行号）
        sheet1.row_dimensions[row_num].height = 33.75  # 设置行高

    

    # 假设以下变量已经定义并赋值：型号, 出厂盘号, 出厂长度, 光缆芯数, 光缆长度
    # 填写表头
    header_rows = [
        ["表四 "+ 地址 +"用户段光纤线路衰减测试记录", "", "", "", "", "", "", "", "", "", "", ""],  # 第2行表头
        ["OTDR:s20C/N          折射率:1.4659          波长:1550 nm          光缆长度:" + 光缆总长 + "km", "", "", "", "", "", "", "", "", "", "", ""],  # 第3行表头数据
        ["光纤序号", "", "损耗","", "光纤序号", "", "损耗", "","光纤序号", "", "损耗", ""],  # 第4行表头数据
        ["", "", "db", "db/km", "", "", "db", "db/km","", "", "db", "db/km"],  # 第5行表头数据
    ]
    start_row = bgs + KS + 1
    for row_data in header_rows:
        for col_index, value in enumerate(row_data, start=1):
            sheet1.cell(row=start_row, column=col_index).value = value
        start_row += 1

    # 页尾
    footer_rows = [
        ["折射波长：1550nm   测试人：                  监理：                日期：          ", "", "", "", "", "", "", "", "", ""]  # 第33行页尾
    ]
    start_row1 = bgs + KS + 21
    for row_data in footer_rows:
        for col_index, value in enumerate(row_data, start=1):
            sheet1.cell(row=start_row1, column=col_index).value = value
        start_row1 += 1

    def 损耗值():
        return round(random.uniform(0.141, 0.131), 3)
    
    def 平均损耗值():
        return round(random.uniform(0.202, 0.191), 3)
    
    for m in range(16):
            # 设置序号
            sheet1.cell(row=bgs + KS + m + 5, column=1).value = m + 1 + 48
            sheet1.cell(row=bgs + KS + m + 5, column=5).value = m + 17 + 48
            sheet1.cell(row=bgs + KS + m + 5, column=9).value = m + 33 + 48
            # 设置A-B
            sheet1.cell(row=bgs + KS + m + 5, column=2).value = "A-B"
            sheet1.cell(row=bgs + KS + m + 5, column=6).value = "A-B"
            sheet1.cell(row=bgs + KS + m + 5, column=10).value = "A-B"
    
    if 光缆芯数 in [48, 96, 144]:
        for m in range(16):
            loss_value = 损耗值()
            loss_value_1 = 平均损耗值()
            # 设置损耗值
            sheet1.cell(row=bgs + KS + m + 5, column=3).value = loss_value
            sheet1.cell(row=bgs + KS + m + 5, column=7).value = loss_value
            sheet1.cell(row=bgs + KS + m + 5, column=11).value = loss_value
            # 设置平均损耗值
            sheet1.cell(row=bgs + KS + m + 5, column=4).value = loss_value_1
            sheet1.cell(row=bgs + KS + m + 5, column=8).value = loss_value_1
            sheet1.cell(row=bgs + KS + m + 5, column=12).value = loss_value_1


    # 定义要合并的单元格区域
    merge_ranges = [
        "A{}:L{}".format(bgs + KS + 1, bgs + KS + 1),  # 合并第一行的A到J列
        "A{}:L{}".format(bgs + KS + 2, bgs + KS + 2),  # 合并第二行的A到J列
        "A{}:B{}".format(bgs + KS + 3, bgs + KS + 4),  # 合并第三行的A、B列
        "E{}:F{}".format(bgs + KS + 3, bgs + KS + 4),  # 合并第三行的A、B列
        "I{}:J{}".format(bgs + KS + 3, bgs + KS + 4),  # 合并第三行的A、B列
        "C{}:D{}".format(bgs + KS + 3, bgs + KS + 3),  # 合并第三行的A、B列
        "G{}:H{}".format(bgs + KS + 3, bgs + KS + 3),  # 合并第三行的A、B列
        "K{}:L{}".format(bgs + KS + 3, bgs + KS + 3),  # 合并第三行的A、B列
        "A{}:L{}".format(bgs + KS + 21, bgs + KS + 21)  # 合并第21行的A到J列
    ]
    
    # 循环遍历合并单元格
    for merge_range in merge_ranges:
        sheet1.merge_cells(merge_range)  # 调用合并单元格的方法，将指定范围的单元格合并

    # 设置A1字体、大小、加粗、自动换行、水平居中
    font = Font(name='宋体', size=14, bold=True)
    sheet1.cell(row=bgs + KS + 1, column=1).font = font
    sheet1.cell(row=bgs + KS + 1, column=1).alignment = Alignment(wrap_text=True,horizontal= 'center', vertical='center')
    


        # 设置 A2:F31 区域的文字上下左右居中
    alignment_center = Alignment(horizontal='center', vertical='center')
    for row in range(bgs + KS + 2, bgs + KS + 22):
        for col in range(1, 13):
            cell = sheet1.cell(row=row, column=col)
            cell.alignment = alignment_center

            # 设置单元格格式，保留3位小数
    number_format = ('0.000')

    # B7:B30 和 E7:E30 范围内的单元格设置格式
    for row in range(bgs + KS + 5, bgs + KS + 21):
        for col in [3, 4]:  # 列C和列D的列号
            cell = sheet1.cell(row=row, column=col)
            cell.number_format = number_format
        for col in [7, 8]:  # 列G和列H的列号
            cell = sheet1.cell(row=row, column=col)
            cell.number_format = number_format
        for col in [11, 12]:  # 列K和列L的列号
            cell = sheet1.cell(row=row, column=col)
            cell.number_format = number_format

    

    # 添加边框线
    border = Border(left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin'))
    for row in range(bgs + KS + 3, bgs + KS + 21):
        for col in range(1, 13):
            cell = sheet1.cell(row=row, column=col)
            cell.border = border

def 第三个表(sheet1, 光缆芯数, KS, 光缆总长, 地址):

    # 光缆总长 = 2000
    # 地址 = "环城北路与云岫北路576光交至北路576光交华凌建材东墙"
    
    bgs = 2 * 21  # 前面表格数量

     # 设置第一行行高
    sheet1.row_dimensions[bgs + KS + 1].height = 45.75
    sheet1.row_dimensions[bgs + KS + 2].height = 21.75
    sheet1.row_dimensions[bgs + KS + 3].height = 21.75
    sheet1.row_dimensions[bgs + KS + 4].height = 21.75
    sheet1.row_dimensions[bgs + KS + 21].height = 21.75

    # 设置 A2 至 A33 行高为 21.75
    for row_num in range(bgs + KS + 5, bgs + KS + 21):  # 从 bgs + KS+2（即A2的行号）遍历到31（即A31的行号）
        sheet1.row_dimensions[row_num].height = 33.75  # 设置行高

    

    # 假设以下变量已经定义并赋值：型号, 出厂盘号, 出厂长度, 光缆芯数, 光缆长度
    # 填写表头
    header_rows = [
        ["表四 "+ 地址 +"用户段光纤线路衰减测试记录", "", "", "", "", "", "", "", "", "", "", ""],  # 第2行表头
        ["OTDR:s20C/N          折射率:1.4659          波长:1550 nm          光缆长度:" + 光缆总长 + "km", "", "", "", "", "", "", "", "", "", "", ""],  # 第3行表头数据
        ["光纤序号", "", "损耗","", "光纤序号", "", "损耗", "","光纤序号", "", "损耗", ""],  # 第4行表头数据
        ["", "", "db", "db/km", "", "", "db", "db/km","", "", "db", "db/km"],  # 第5行表头数据
    ]
    start_row = bgs + KS + 1
    for row_data in header_rows:
        for col_index, value in enumerate(row_data, start=1):
            sheet1.cell(row=start_row, column=col_index).value = value
        start_row += 1

    # 页尾
    footer_rows = [
        ["折射波长：1550nm   测试人：                  监理：                日期：          ", "", "", "", "", "", "", "", "", ""]  # 第33行页尾
    ]
    start_row1 = bgs + KS + 21
    for row_data in footer_rows:
        for col_index, value in enumerate(row_data, start=1):
            sheet1.cell(row=start_row1, column=col_index).value = value
        start_row1 += 1

    def 损耗值():
        return round(random.uniform(0.141, 0.131), 3)
    
    def 平均损耗值():
        return round(random.uniform(0.202, 0.191), 3)
    
    for m in range(16):
            # 设置序号
            sheet1.cell(row=bgs + KS + m + 5, column=1).value = m + 1 + 96
            sheet1.cell(row=bgs + KS + m + 5, column=5).value = m + 17 + 96
            sheet1.cell(row=bgs + KS + m + 5, column=9).value = m + 33 + 96
            # 设置A-B
            sheet1.cell(row=bgs + KS + m + 5, column=2).value = "A-B"
            sheet1.cell(row=bgs + KS + m + 5, column=6).value = "A-B"
            sheet1.cell(row=bgs + KS + m + 5, column=10).value = "A-B"
    
    if 光缆芯数 in [48, 96, 144]:
        for m in range(16):
            loss_value = 损耗值()
            loss_value_1 = 平均损耗值()
            # 设置损耗值
            sheet1.cell(row=bgs + KS + m + 5, column=3).value = loss_value
            sheet1.cell(row=bgs + KS + m + 5, column=7).value = loss_value
            sheet1.cell(row=bgs + KS + m + 5, column=11).value = loss_value
            # 设置平均损耗值
            sheet1.cell(row=bgs + KS + m + 5, column=4).value = loss_value_1
            sheet1.cell(row=bgs + KS + m + 5, column=8).value = loss_value_1
            sheet1.cell(row=bgs + KS + m + 5, column=12).value = loss_value_1

       

    # 定义要合并的单元格区域
    merge_ranges = [
        "A{}:L{}".format(bgs + KS + 1, bgs + KS + 1),  # 合并第一行的A到J列
        "A{}:L{}".format(bgs + KS + 2, bgs + KS + 2),  # 合并第二行的A到J列
        "A{}:B{}".format(bgs + KS + 3, bgs + KS + 4),  # 合并第三行的A、B列
        "E{}:F{}".format(bgs + KS + 3, bgs + KS + 4),  # 合并第三行的A、B列
        "I{}:J{}".format(bgs + KS + 3, bgs + KS + 4),  # 合并第三行的A、B列
        "C{}:D{}".format(bgs + KS + 3, bgs + KS + 3),  # 合并第三行的A、B列
        "G{}:H{}".format(bgs + KS + 3, bgs + KS + 3),  # 合并第三行的A、B列
        "K{}:L{}".format(bgs + KS + 3, bgs + KS + 3),  # 合并第三行的A、B列
        "A{}:L{}".format(bgs + KS + 21, bgs + KS + 21)  # 合并第21行的A到J列
    ]
    
    # 循环遍历合并单元格
    for merge_range in merge_ranges:
        sheet1.merge_cells(merge_range)  # 调用合并单元格的方法，将指定范围的单元格合并

    # 设置A1字体、大小、加粗、自动换行、水平居中
    font = Font(name='宋体', size=14, bold=True)
    sheet1.cell(row=bgs + KS + 1, column=1).font = font
    sheet1.cell(row=bgs + KS + 1, column=1).alignment = Alignment(wrap_text=True,horizontal= 'center', vertical='center')
    


        # 设置 A2:F31 区域的文字上下左右居中
    alignment_center = Alignment(horizontal='center', vertical='center')
    for row in range(bgs + KS + 2, bgs + KS + 22):
        for col in range(1, 13):
            cell = sheet1.cell(row=row, column=col)
            cell.alignment = alignment_center

            # 设置单元格格式，保留3位小数
    number_format = ('0.000')

    # B7:B30 和 E7:E30 范围内的单元格设置格式
    for row in range(bgs + KS + 5, bgs + KS + 21):
        for col in [3, 4]:  # 列C和列D的列号
            cell = sheet1.cell(row=row, column=col)
            cell.number_format = number_format
        for col in [7, 8]:  # 列G和列H的列号
            cell = sheet1.cell(row=row, column=col)
            cell.number_format = number_format
        for col in [11, 12]:  # 列K和列L的列号
            cell = sheet1.cell(row=row, column=col)
            cell.number_format = number_format

    

    # 添加边框线
    border = Border(left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin'))
    for row in range(bgs + KS + 3, bgs + KS + 21):
        for col in range(1, 13):
            cell = sheet1.cell(row=row, column=col)
            cell.border = border


def glcs_b3(sheet1, 光缆芯数, 光缆总长, 地址):

    row = sheet1.max_row
    if row == 1:
        max_row = 0
    else:
        max_row = row

    if 光缆芯数 in [4, 8, 12, 24, 48]:
        # 执行程序1
        第一个表(sheet1, 光缆芯数,  max_row, 光缆总长, 地址)  # 调用第一个表

    elif 光缆芯数 == 96:
        # 执行程序2
        第一个表(sheet1, 光缆芯数,  max_row, 光缆总长, 地址)  # 调用第一个表
        第二个表(sheet1, 光缆芯数,  max_row, 光缆总长, 地址)  # 调用第一个表

        # 在这里编写程序2的具体代码
    elif 光缆芯数 == 144:
        # 执行程序3
        第一个表(sheet1, 光缆芯数,  max_row, 光缆总长, 地址)  # 调用第一个表
        第二个表(sheet1, 光缆芯数,  max_row, 光缆总长, 地址)  # 调用第二个表
        第三个表(sheet1, 光缆芯数,  max_row, 光缆总长, 地址)  # 调用第三个表

        # 在这里编写程序3的具体代码
    else:
        # 如果A的值不是上述任何一个，执行此处的代码（可选）
        print("A的值不是预期的，不执行任何程序")

    # 名称 = 型号 + "-" + str(光缆芯数) + "接头损耗表.xlsx"
    # wb.save(名称)

    # 每张表 21 行、12 列；让所有表各占一页
    设置每页一个表(sheet1, 每个表行数=21, 每个表列数=12, 横向=True)
