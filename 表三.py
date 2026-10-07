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


def 第一个表(sheet1, 光缆芯数, KS, 光缆总长, 接头数, 地址):

    # 光缆总长 = 2000
    # 地址 = "环城北路与云岫北路576光交至北路576光交华凌建材东墙"
    接头编号 = str(接头数) + "号"

    # 设置第一行行高
    sheet1.row_dimensions[KS + 1].height = 45.75

    # 设置 A2 至 A33 行高为 21.75
    for row_num in range(KS + 2, KS + 31):  # 从 KS+2（即A2的行号）遍历到31（即A31的行号）
        sheet1.row_dimensions[row_num].height = 21.75  # 设置行高

    

    # 假设以下变量已经定义并赋值：型号, 出厂盘号, 出厂长度, 光缆芯数, 光缆长度
    # 填写表头
    header_rows = [
        ["表三 "+地址+"用户段接头衰耗测试表", "", "", "", "", "", "", "", "", ""],  # 第2行表头
        ["熔纤机:ALK-80          OTDR:s20C/N          折射率:1.4659         波长:1550 nm     ", "", "", "", "", "", "", "", "", ""],  # 第3行表头数据
        ["接头编号:", "", 接头编号, "", "", "纤长:", "", 光缆总长, "", ""],  # 第4行表头数据
        ["光纤序号", "衰耗（dB）", "", "", "", "光纤序号", "衰耗（dB）", "", "", ""],  # 第5行表头数据
        ["", "A-B", "B-A", "合计", "平均", "", "A-B", "B-A", "合计", "平均"],  # 第6行表头数据
    ]
    start_row = KS + 1
    for row_data in header_rows:
        for col_index, value in enumerate(row_data, start=1):
            sheet1.cell(row=start_row, column=col_index).value = value
        start_row += 1

    # 页尾
    footer_rows = [
        ["熔接人：              测试人：              监理：              日期：          ", "", "", "", "", "", "", "", "", ""]  # 第33行页尾
    ]
    start_row1 = KS + 30
    for row_data in footer_rows:
        for col_index, value in enumerate(row_data, start=1):
            sheet1.cell(row=start_row1, column=col_index).value = value
        start_row1 += 1

        def 损耗值():
            return round(random.uniform(0.021, 0.002), 3)
        
        
        if 光缆芯数 in [48, 96, 144]:
            for m in range(24):
                loss_value = 损耗值()
                loss_value_1 = 损耗值()
                sheet1.cell(row=KS + m + 6, column=1).value = m + 1
                sheet1.cell(row=KS + m + 6, column=6).value = m + 25
                sheet1.cell(row=KS + m + 6, column=2).value = loss_value
                sheet1.cell(row=KS + m + 6, column=5).value = loss_value
                sheet1.cell(row=KS + m + 6, column=7).value = loss_value_1
                sheet1.cell(row=KS + m + 6, column=10).value = loss_value_1

        else:
            for i in [4, 8, 12, 24]:
                if 光缆芯数 == i:
                    for m in range(i):
                        loss_value = 损耗值()
                        sheet1.cell(row=KS + m + 6, column=1).value = m + 1
                        sheet1.cell(row=KS + m + 6, column=2).value = loss_value
                        sheet1.cell(row=KS + m + 6, column=5).value = loss_value
                    break
            # else:
            # print("Invalid input")

    # 定义要合并的单元格区域
    merge_ranges = [
        "A{}:J{}".format(KS + 1, KS + 1),  # 合并第一行的A到J列
        "A{}:J{}".format(KS + 2, KS + 2),  # 合并第二行的A到J列
        "A{}:B{}".format(KS + 3, KS + 3),  # 合并第三行的A、B列
        "C{}:E{}".format(KS + 3, KS + 3),  # 合并第三行的C到E列
        "F{}:G{}".format(KS + 3, KS + 3),  # 合并第三行的F、G列
        "H{}:J{}".format(KS + 3, KS + 3),  # 合并第三行的H到J列
        "A{}:A{}".format(KS + 4, KS + 5),  # 合并第四行到第五行的A列
        "F{}:F{}".format(KS + 4, KS + 5),  # 合并第四行到第五行的F列
        "B{}:E{}".format(KS + 4, KS + 4),  # 合并第四行的B到E列
        "G{}:J{}".format(KS + 4, KS + 4),  # 合并第四行的G到J列
        "A{}:J{}".format(KS + 30, KS + 30)  # 合并第30行的A到J列
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
    for row in range(KS + 2, KS + 31):
        for col in range(1, 11):
            cell = sheet1.cell(row=row, column=col)
            cell.alignment = alignment_center

            # 设置单元格格式，保留3位小数
    number_format = ('0.000')

    # B7:B30 和 E7:E30 范围内的单元格设置格式
    for row in range(KS + 6, KS + 30):
        for col in [2, 5]:  # 列B和列E的列号
            cell = sheet1.cell(row=row, column=col)
            cell.number_format = number_format
        for col in [7, 10]:  # 列B和列E的列号
            cell = sheet1.cell(row=row, column=col)
            cell.number_format = number_format

    

    # 添加边框线
    border = Border(left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin'))
    for row in range(KS + 3, KS + 30):
        for col in range(1, 11):
            cell = sheet1.cell(row=row, column=col)
            cell.border = border

def 第二个表(sheet1, 光缆芯数, KS, 光缆总长, 接头数, 地址):

    # 光缆总长 = 2000
    # 地址 = "环城北路与云岫北路576光交至北路576光交华凌建材东墙"
    接头编号 = str(接头数) + "号"

    bgs = 1 * 30  # 前面表格数量

    # 设置第一行行高
    sheet1.row_dimensions[bgs + KS + 1].height = 45.75

    # 设置 A2 至 A33 行高为 21.75
    for row_num in range(bgs + KS + 2, bgs + KS + 31):  # 从 bgs + KS+2（即A2的行号）遍历到31（即A31的行号）
        sheet1.row_dimensions[row_num].height = 21.75  # 设置行高

    

    # 假设以下变量已经定义并赋值：型号, 出厂盘号, 出厂长度, 光缆芯数, 光缆长度
    # 填写表头
    header_rows = [
        ["表三 "+地址+"用户段接头衰耗测试表", "", "", "", "", "", "", "", "", ""],  # 第2行表头
        ["熔纤机:ALK-80          OTDR:s20C/N          折射率:1.4659         波长:1550 nm     ", "", "", "", "", "", "", "", "", ""],  # 第3行表头数据
        ["接头编号:", "", 接头编号, "", "", "纤长:", "", 光缆总长, "", ""],  # 第4行表头数据
        ["光纤序号", "衰耗（dB）", "", "", "", "光纤序号", "衰耗（dB）", "", "", ""],  # 第5行表头数据
        ["", "A-B", "B-A", "合计", "平均", "", "A-B", "B-A", "合计", "平均"],  # 第6行表头数据
    ]
    start_row = bgs + KS + 1
    for row_data in header_rows:
        for col_index, value in enumerate(row_data, start=1):
            sheet1.cell(row=start_row, column=col_index).value = value
        start_row += 1

    # 页尾
    footer_rows = [
        ["熔接人：              测试人：              监理：              日期：          ", "", "", "", "", "", "", "", "", ""]  # 第33行页尾
    ]
    start_row1 = bgs + KS + 30
    for row_data in footer_rows:
        for col_index, value in enumerate(row_data, start=1):
            sheet1.cell(row=start_row1, column=col_index).value = value
        start_row1 += 1

        def 损耗值():
            return round(random.uniform(0.021, 0.002), 3)
        
        
        if 光缆芯数 in [48, 96, 144]:
            for m in range(24):
                loss_value = 损耗值()
                loss_value_1 = 损耗值()
                sheet1.cell(row=bgs + KS + m + 6, column=1).value = m + 1 + 48
                sheet1.cell(row=bgs + KS + m + 6, column=6).value = m + 25 + 48
                sheet1.cell(row=bgs + KS + m + 6, column=2).value = loss_value
                sheet1.cell(row=bgs + KS + m + 6, column=5).value = loss_value
                sheet1.cell(row=bgs + KS + m + 6, column=7).value = loss_value_1
                sheet1.cell(row=bgs + KS + m + 6, column=10).value = loss_value_1


    # 定义要合并的单元格区域
    merge_ranges = [
        "A{}:J{}".format(bgs + KS + 1, bgs + KS + 1),  # 合并第一行的A到J列
        "A{}:J{}".format(bgs + KS + 2, bgs + KS + 2),  # 合并第二行的A到J列
        "A{}:B{}".format(bgs + KS + 3, bgs + KS + 3),  # 合并第三行的A、B列
        "C{}:E{}".format(bgs + KS + 3, bgs + KS + 3),  # 合并第三行的C到E列
        "F{}:G{}".format(bgs + KS + 3, bgs + KS + 3),  # 合并第三行的F、G列
        "H{}:J{}".format(bgs + KS + 3, bgs + KS + 3),  # 合并第三行的H到J列
        "A{}:A{}".format(bgs + KS + 4, bgs + KS + 5),  # 合并第四行到第五行的A列
        "F{}:F{}".format(bgs + KS + 4, bgs + KS + 5),  # 合并第四行到第五行的F列
        "B{}:E{}".format(bgs + KS + 4, bgs + KS + 4),  # 合并第四行的B到E列
        "G{}:J{}".format(bgs + KS + 4, bgs + KS + 4),  # 合并第四行的G到J列
        "A{}:J{}".format(bgs + KS + 30, bgs + KS + 30)  # 合并第30行的A到J列
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
    for row in range(bgs + KS + 2, bgs + KS + 31):
        for col in range(1, 11):
            cell = sheet1.cell(row=row, column=col)
            cell.alignment = alignment_center

            # 设置单元格格式，保留3位小数
    number_format = ('0.000')

    # B7:B30 和 E7:E30 范围内的单元格设置格式
    for row in range(bgs + KS + 6, bgs + KS + 30):
        for col in [2, 5]:  # 列B和列E的列号
            cell = sheet1.cell(row=row, column=col)
            cell.number_format = number_format
        for col in [7, 10]:  # 列B和列E的列号
            cell = sheet1.cell(row=row, column=col)
            cell.number_format = number_format

    

    # 添加边框线
    border = Border(left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin'))
    for row in range(bgs + KS + 3, bgs + KS + 30):
        for col in range(1, 11):
            cell = sheet1.cell(row=row, column=col)
            cell.border = border

def 第三个表(sheet1, 光缆芯数, KS, 光缆总长, 接头数, 地址):

    # 光缆总长 = 2000
    # 地址 = "环城北路与云岫北路576光交至北路576光交华凌建材东墙"
    接头编号 = str(接头数) + "号"

    bgs = 2 * 30  # 前面表格数量

    # 设置第一行行高
    sheet1.row_dimensions[bgs + KS + 1].height = 45.75

    # 设置 A2 至 A33 行高为 21.75
    for row_num in range(bgs + KS + 2, bgs + KS + 31):  # 从 bgs + KS+2（即A2的行号）遍历到31（即A31的行号）
        sheet1.row_dimensions[row_num].height = 21.75  # 设置行高

    

    # 假设以下变量已经定义并赋值：型号, 出厂盘号, 出厂长度, 光缆芯数, 光缆长度
    # 填写表头
    header_rows = [
        ["表三 "+地址+"用户段接头衰耗测试表", "", "", "", "", "", "", "", "", ""],  # 第2行表头
        ["熔纤机:ALK-80          OTDR:s20C/N          折射率:1.4659         波长:1550 nm     ", "", "", "", "", "", "", "", "", ""],  # 第3行表头数据
        ["接头编号:", "", 接头编号, "", "", "纤长:", "", 光缆总长, "", ""],  # 第4行表头数据
        ["光纤序号", "衰耗（dB）", "", "", "", "光纤序号", "衰耗（dB）", "", "", ""],  # 第5行表头数据
        ["", "A-B", "B-A", "合计", "平均", "", "A-B", "B-A", "合计", "平均"],  # 第6行表头数据
    ]
    start_row = bgs + KS + 1
    for row_data in header_rows:
        for col_index, value in enumerate(row_data, start=1):
            sheet1.cell(row=start_row, column=col_index).value = value
        start_row += 1

    # 页尾
    footer_rows = [
        ["熔接人：              测试人：              监理：              日期：          ", "", "", "", "", "", "", "", "", ""]  # 第33行页尾
    ]
    start_row1 = bgs + KS + 30
    for row_data in footer_rows:
        for col_index, value in enumerate(row_data, start=1):
            sheet1.cell(row=start_row1, column=col_index).value = value
        start_row1 += 1

        def 损耗值():
            return round(random.uniform(0.021, 0.002), 3)
        
        
        if 光缆芯数 in [48, 96, 144]:
            for m in range(24):
                loss_value = 损耗值()
                loss_value_1 = 损耗值()
                sheet1.cell(row=bgs + KS + m + 6, column=1).value = m + 1 + 96
                sheet1.cell(row=bgs + KS + m + 6, column=6).value = m + 25 + 96
                sheet1.cell(row=bgs + KS + m + 6, column=2).value = loss_value
                sheet1.cell(row=bgs + KS + m + 6, column=5).value = loss_value
                sheet1.cell(row=bgs + KS + m + 6, column=7).value = loss_value_1
                sheet1.cell(row=bgs + KS + m + 6, column=10).value = loss_value_1


    # 定义要合并的单元格区域
    merge_ranges = [
        "A{}:J{}".format(bgs + KS + 1, bgs + KS + 1),  # 合并第一行的A到J列
        "A{}:J{}".format(bgs + KS + 2, bgs + KS + 2),  # 合并第二行的A到J列
        "A{}:B{}".format(bgs + KS + 3, bgs + KS + 3),  # 合并第三行的A、B列
        "C{}:E{}".format(bgs + KS + 3, bgs + KS + 3),  # 合并第三行的C到E列
        "F{}:G{}".format(bgs + KS + 3, bgs + KS + 3),  # 合并第三行的F、G列
        "H{}:J{}".format(bgs + KS + 3, bgs + KS + 3),  # 合并第三行的H到J列
        "A{}:A{}".format(bgs + KS + 4, bgs + KS + 5),  # 合并第四行到第五行的A列
        "F{}:F{}".format(bgs + KS + 4, bgs + KS + 5),  # 合并第四行到第五行的F列
        "B{}:E{}".format(bgs + KS + 4, bgs + KS + 4),  # 合并第四行的B到E列
        "G{}:J{}".format(bgs + KS + 4, bgs + KS + 4),  # 合并第四行的G到J列
        "A{}:J{}".format(bgs + KS + 30, bgs + KS + 30)  # 合并第30行的A到J列
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
    for row in range(bgs + KS + 2, bgs + KS + 31):
        for col in range(1, 11):
            cell = sheet1.cell(row=row, column=col)
            cell.alignment = alignment_center

            # 设置单元格格式，保留3位小数
    number_format = ('0.000')

    # B7:B30 和 E7:E30 范围内的单元格设置格式
    for row in range(bgs + KS + 6, bgs + KS + 30):
        for col in [2, 5]:  # 列B和列E的列号
            cell = sheet1.cell(row=row, column=col)
            cell.number_format = number_format
        for col in [7, 10]:  # 列B和列E的列号
            cell = sheet1.cell(row=row, column=col)
            cell.number_format = number_format

    

    # 添加边框线
    border = Border(left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin'))
    for row in range(bgs + KS + 3, bgs + KS + 30):
        for col in range(1, 11):
            cell = sheet1.cell(row=row, column=col)
            cell.border = border


def 表三(sheet1, 光缆芯数, 光缆总长, 接头数, 地址):

    for i in range(接头数):
        jts = i + 1

        row = sheet1.max_row
        if row == 1:
            max_row = 0
        else:
            max_row = row

        if 光缆芯数 in [4, 8, 12, 24, 48]:
            # 执行程序1
            第一个表(sheet1, 光缆芯数,  max_row, 光缆总长, jts, 地址)  # 调用第一个表

        elif 光缆芯数 == 96:
            # 执行程序2
            第一个表(sheet1, 光缆芯数,  max_row, 光缆总长, jts, 地址)  # 调用第一个表
            第二个表(sheet1, 光缆芯数,  max_row, 光缆总长, jts, 地址)  # 调用第一个表

            # 在这里编写程序2的具体代码
        elif 光缆芯数 == 144:
            # 执行程序3
            第一个表(sheet1, 光缆芯数,  max_row, 光缆总长, jts, 地址)  # 调用第一个表
            第二个表(sheet1, 光缆芯数,  max_row, 光缆总长, jts, 地址)  # 调用第二个表
            第三个表(sheet1, 光缆芯数,  max_row, 光缆总长, jts, 地址)  # 调用第三个表

            # 在这里编写程序3的具体代码
        else:
            # 如果A的值不是上述任何一个，执行此处的代码（可选）
            print("A的值不是预期的，不执行任何程序")

        # 名称 = 型号 + "-" + str(光缆芯数) + "接头损耗表.xlsx"
        # wb.save(名称)

        jts += 1  # 接头序号加1

    # 每张表 30 行、10 列；让所有表各占一页
    设置每页一个表(sheet1, 每个表行数=30, 每个表列数=10, 横向=True)



# # 读取excel文件
# df = pd.read_excel('批量打印模板（不用打印1）.xlsx')

# # 去除'接头数'列为空的行
# df = df.dropna(subset=['接头数'])

# # 计算df列表的总行数
# total_rows = df.shape[0]

# # 光缆芯数 = 12
# # 页数 = 2
# # 光缆总长 = 1000
# # 接头数 = 3
# # 地址 = "环城北路与云岫北路576光交至北路576光交华凌建材东墙"


# # 创建一个新的工作簿
# wb = openpyxl.Workbook()

# # 获取默认的活动工作表
# sheet1 = wb.active


# # 设置 A1 至 J1 列的宽度为 6.75
# for col in sheet1.iter_cols(min_col=1, max_col=10):
#     for cell in col:
#         sheet1.column_dimensions[cell.column_letter].width = 8.4


# for i in range(total_rows):
#     # 输出第一行第一列和第一行第三列的数据
#     地址 = df.iloc[i, 1]
#     光缆芯数 = int(df.iloc[i, 2][:-1])
#     光缆总长 = df.iloc[i, 3]
#     接头数 = int(df.iloc[i, 4])

#     表三(sheet1, 光缆芯数, 光缆总长, 接头数, 地址)


# # 保存工作簿
# 名称 = "表三接头损耗表.xlsx"
# wb.save(名称)