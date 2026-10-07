# from openpyxl import load_workbook
import openpyxl
import random
from openpyxl.styles import Border, Side, Alignment, Font
from openpyxl.worksheet.pagebreak import Break
from openpyxl.worksheet.properties import PageSetupProperties
from datetime import datetime
from dateutil.relativedelta import relativedelta
import calendar


def 设置每页一个表(sheet, 每个表行数, 每个表列数, 横向=False):
    """
    根据每个表所占行数，自动在行边界插入分页符，
    并设置 A4 纸张、缩放适应一页宽一页高。
    """
    max_row = sheet.max_row
    max_col = 每个表列数
    # 清掉旧的分页符
    sheet.row_breaks.brk = []
    # 在每个表的结束行（最后一张表除外）后面插入分页符
    if 每个表行数 > 0:
        表数量 = max_row // 每个表行数
        for i in range(1, 表数量):
            sheet.row_breaks.append(Break(id=i * 每个表行数))

    # 页面设置：A4，方向；不使用 fitToPage 缩放，让内容按 100% 显示
    # 分页符会保证每张表独占一页
    sheet.page_setup.paperSize = sheet.PAPERSIZE_A4
    sheet.page_setup.orientation = 'landscape' if 横向 else 'portrait'
    sheet.page_setup.scale = 100  # 100% 缩放，不压缩
    sheet.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=False)

    # 打印区域：覆盖到最大行
    from openpyxl.utils import get_column_letter
    last_col_letter = get_column_letter(max_col)
    sheet.print_area = 'A1:{}{}'.format(last_col_letter, max_row)

# 型号 = "GTYA"
# 光缆芯数 = 12
# 页数 = 2

def dpcs_zcx(型号,光缆芯数,页数):

    # 创建一个新的工作簿
    wb = openpyxl.Workbook()

    # 获取默认的活动工作表
    sheet1 = wb.active

    # 填入数据
    # 型号 = 型号
    # 光缆芯数 = 光缆芯数
    # 页数 = 页数


    # 设置 A1 至 F1 列的宽度为 13.75
    for col in sheet1.iter_cols(min_col=1, max_col=6):
        for cell in col:
            sheet1.column_dimensions[cell.column_letter].width = 14.37


    if 光缆芯数 in [4, 8, 12, 24, 48]:
        # 执行程序1
        my_function48(sheet1, 型号, 光缆芯数, 页数)

    elif 光缆芯数 == 96:
        # 执行程序2
        my_function96(sheet1, 型号, 光缆芯数, 页数)

        # 在这里编写程序2的具体代码
    elif 光缆芯数 == 144:
        # 执行程序3
        my_function144(sheet1, 型号, 光缆芯数, 页数)
        # 在这里编写程序3的具体代码
    else:
        # 如果A的值不是上述任何一个，执行此处的代码（可选）
        print("A的值不是预期的，不执行任何程序")

    # 每个表占 33 行，6 列；让每张表打印时独占一页
    设置每页一个表(sheet1, 每个表行数=33, 每个表列数=6, 横向=False)

    名称 = 型号 + "-" + str(光缆芯数) + "B1单盘测试.xlsx"
    wb.save(名称)



def 第一个表(sheet1, 型号, 光缆芯数, 出厂盘号, KS, 测试光纤长度):
    出厂长度 = 2.000
    光缆长度 = 2.000

    # 设置第一行行高
    sheet1.row_dimensions[KS + 1].height = 36

    # 设置 A2 至 A33 行高为 20.25
    for row_num in range(KS + 2, KS + 34):  # 从 KS+2（即A2的行号）遍历到33（即A33的行号）
        sheet1.row_dimensions[row_num].height = 20.25  # 设置行高

    # 设置字体
    font = Font(name='宋体', size=16, bold=True)
    sheet1.cell(row=KS + 1, column=1).font = font

    # 假设以下变量已经定义并赋值：型号, 出厂盘号, 出厂长度, 光缆芯数, 光缆长度
    # 填写表头
    header_rows = [
        ["附表2 光缆单盘检验测试记录表", "", "", "", "", ""],  # 第2行表头
        ["光缆型号", 型号, "出厂盘号", 出厂盘号, "出厂长度", 出厂长度],  # 第3行表头数据
        ["光缆芯数", str(光缆芯数) + "B1", "光缆长度", 光缆长度, "外观", "完整"],  # 第4行表头数据
        ["测试仪表", "AQ7260", "单位", "km", "≤0.22dB/km", ""],  # 第5行表头数据
        ["纤芯序号", "1550nm（折射率1.4680）", "", "纤芯序号", "1550nm（折射率1.4680）", ""],  # 第6行表头数据
        ["", "测试衰耗值", "测试光纤长度", "", "测试衰耗值", "测试光纤长度"],  # 第7行表头数据
    ]
    start_row = KS + 1
    for row_data in header_rows:
        for col_index, value in enumerate(row_data, start=1):
            sheet1.cell(row=start_row, column=col_index).value = value
        start_row += 1

    # 页尾
    footer_rows = [
        ["结论", "", "", "", "", ""],
        ["测试人员：", "", "", "监理人员：", "", ""],  # 第32行页尾
        ["日期： ", "", "", "日期： ", "", ""]  # 第33行页尾
    ]
    start_row1 = KS + 31
    for row_data in footer_rows:
        for col_index, value in enumerate(row_data, start=1):
            sheet1.cell(row=start_row1, column=col_index).value = value
        start_row1 += 1

        def 测试衰耗值():
            return round(random.uniform(0.185, 0.2), 3)

        # 光缆芯数 = 8
        if 光缆芯数 in [48, 96, 144]:
            for m in range(24):
                sheet1.cell(row=KS + m + 7, column=1).value = m + 1
                sheet1.cell(row=KS + m + 7, column=4).value = m + 25
                sheet1.cell(row=KS + m + 7, column=2).value = 测试衰耗值()
                sheet1.cell(row=KS + m + 7, column=5).value = 测试衰耗值()
                sheet1.cell(row=KS + m + 7, column=3).value = 测试光纤长度
                sheet1.cell(row=KS + m + 7, column=6).value = 测试光纤长度

        else:
            for i in [4, 8, 12, 24]:
                if 光缆芯数 == i:
                    for m in range(i):
                        sheet1.cell(row=KS + m + 7, column=1).value = m + 1
                        sheet1.cell(row=KS + m + 7, column=2).value = 测试衰耗值()
                        sheet1.cell(row=KS + m + 7, column=3).value = 测试光纤长度
                    break
            # else:
            # print("Invalid input")

    # 定义要合并的单元格区域
    merge_ranges = [
        "A{}:F{}".format(KS + 1, KS + 1),
        "B{}:C{}".format(KS + 5, KS + 5),
        "E{}:F{}".format(KS + 4, KS + 4),
        "E{}:F{}".format(KS + 5, KS + 5),
        "A{}:A{}".format(KS + 5, KS + 6),
        "D{}:D{}".format(KS + 5, KS + 6),
        "B{}:F{}".format(KS + 31, KS + 31)
    ]

    # 循环遍历合并单元格
    for merge_range in merge_ranges:
        sheet1.merge_cells(merge_range)

        # 设置 A1:F31 区域的文字上下左右居中
    alignment_center = Alignment(horizontal='center', vertical='center')
    for row in range(KS + 1, KS + 34):
        for col in range(1, 7):
            cell = sheet1.cell(row=row, column=col)
            cell.alignment = alignment_center

            # 设置单元格格式，保留3位小数
    number_format = ('0.000')

    sheet1.cell(KS + 2, 6).number_format = number_format
    sheet1.cell(KS + 3, 4).number_format = number_format

    # B7:B30 和 E7:E30 范围内的单元格设置格式
    for row in range(KS + 7, KS + 31):
        for col in [2, 3, 5, 6]:  # 列B和列E的列号
            cell = sheet1.cell(row=row, column=col)
            cell.number_format = number_format

            # 添加边框线
    border = Border(left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin'))
    for row in range(KS + 2, KS + 32):
        for col in range(1, 7):
            cell = sheet1.cell(row=row, column=col)
            cell.border = border


def 第二个表(sheet1, 型号, 光缆芯数, 出厂盘号, KS, 测试光纤长度):
    出厂长度 = 2.000
    光缆长度 = 2.000
    bgs = 1 * 33  # 前面表格数量

    # 设置第一行行高
    sheet1.row_dimensions[bgs + KS + 1].height = 36

    # 设置 A2 至 A33 行高为 20.25
    for row_num in range(bgs + KS + 2, bgs + KS + 34):  # 从 bgs + KS+2（即A2的行号）遍历到33（即A33的行号）
        sheet1.row_dimensions[row_num].height = 20.25  # 设置行高

    # 设置字体
    font = Font(name='宋体', size=16, bold=True)
    sheet1.cell(row=bgs + KS + 1, column=1).font = font

    # 假设以下变量已经定义并赋值：型号, 出厂盘号, 出厂长度, 光缆芯数, 光缆长度
    # 填写表头
    header_rows = [
        ["附表2 光缆单盘检验测试记录表", "", "", "", "", ""],  # 第2行表头
        ["光缆型号", 型号, "出厂盘号", 出厂盘号, "出厂长度", 出厂长度],  # 第3行表头数据
        ["光缆芯数", str(光缆芯数) + "B1", "光缆长度", 光缆长度, "外观", "完整"],  # 第4行表头数据
        ["测试仪表", "AQ7260", "单位", "km", "≤0.22dB/km", ""],  # 第5行表头数据
        ["纤芯序号", "1550nm（折射率1.4680）", "", "纤芯序号", "1550nm（折射率1.4680）", ""],  # 第6行表头数据
        ["", "测试衰耗值", "测试光纤长度", "", "测试衰耗值", "测试光纤长度"],  # 第7行表头数据
    ]
    start_row = bgs + KS + 1
    for row_data in header_rows:
        for col_index, value in enumerate(row_data, start=1):
            sheet1.cell(row=start_row, column=col_index).value = value
        start_row += 1

    # 页尾
    footer_rows = [
        ["结论", "", "", "", "", ""],
        ["测试人员：", "", "", "监理人员：", "", ""],  # 第32行页尾
        ["日期： ", "", "", "日期： ", "", ""]  # 第33行页尾
    ]
    start_row1 = bgs + KS + 31
    for row_data in footer_rows:
        for col_index, value in enumerate(row_data, start=1):
            sheet1.cell(row=start_row1, column=col_index).value = value
        start_row1 += 1

        # 填入对应数据
        # 生成一个2000至2020之间的随机数，数字取整

        def 测试衰耗值():
            return round(random.uniform(0.185, 0.2), 3)

        # 光缆芯数 = 8
        if 光缆芯数 in [48, 96, 144]:
            for m in range(24):
                sheet1.cell(row=bgs + KS + m + 7, column=1).value = m + 1 + 48
                sheet1.cell(row=bgs + KS + m + 7, column=4).value = m + 25 + 48
                sheet1.cell(row=bgs + KS + m + 7, column=2).value = 测试衰耗值()
                sheet1.cell(row=bgs + KS + m + 7, column=5).value = 测试衰耗值()
                sheet1.cell(row=bgs + KS + m + 7, column=3).value = 测试光纤长度
                sheet1.cell(row=bgs + KS + m + 7, column=6).value = 测试光纤长度

        else:
            for i in [4, 8, 12, 24]:
                if 光缆芯数 == i:
                    for m in range(i):
                        sheet1.cell(row=bgs + KS + m + 7, column=1).value = m + 1
                        sheet1.cell(row=bgs + KS + m + 7, column=2).value = 测试衰耗值()
                        sheet1.cell(row=bgs + KS + m + 7, column=3).value = 测试光纤长度
                    break
            # else:
            # print("Invalid input")

    # 定义要合并的单元格区域
    merge_ranges = [
        "A{}:F{}".format(bgs + KS + 1, bgs + KS + 1),
        "B{}:C{}".format(bgs + KS + 5, bgs + KS + 5),
        "E{}:F{}".format(bgs + KS + 4, bgs + KS + 4),
        "E{}:F{}".format(bgs + KS + 5, bgs + KS + 5),
        "A{}:A{}".format(bgs + KS + 5, bgs + KS + 6),
        "D{}:D{}".format(bgs + KS + 5, bgs + KS + 6),
        "B{}:F{}".format(bgs + KS + 31, bgs + KS + 31)
    ]

    # 循环遍历合并单元格
    for merge_range in merge_ranges:
        sheet1.merge_cells(merge_range)

        # 设置 A1:F31 区域的文字上下左右居中
    alignment_center = Alignment(horizontal='center', vertical='center')
    for row in range(bgs + KS + 1, bgs + KS + 34):
        for col in range(1, 7):
            cell = sheet1.cell(row=row, column=col)
            cell.alignment = alignment_center

            # 设置单元格格式，保留3位小数
    number_format = ('0.000')

    sheet1.cell(bgs + KS + 2, 6).number_format = number_format
    sheet1.cell(bgs + KS + 3, 4).number_format = number_format

    # B7:B30 和 E7:E30 范围内的单元格设置格式
    for row in range(bgs + KS + 7, bgs + KS + 31):
        for col in [2, 3, 5, 6]:  # 列B,C和列E,F的列号
            cell = sheet1.cell(row=row, column=col)
            cell.number_format = number_format

            # 添加边框线
    border = Border(left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin'))
    for row in range(bgs + KS + 2, bgs + KS + 32):
        for col in range(1, 7):
            cell = sheet1.cell(row=row, column=col)
            cell.border = border


def 第三个表(sheet1, 型号, 光缆芯数, 出厂盘号, KS, 测试光纤长度):
    出厂长度 = 2.000
    光缆长度 = 2.000
    bgs = 2 * 33  # 前面表格数量

    # 设置第一行行高
    sheet1.row_dimensions[bgs + KS + 1].height = 36

    # 设置 A2 至 A33 行高为 20.25
    for row_num in range(bgs + KS + 2, bgs + KS + 34):  # 从 bgs + KS+2（即A2的行号）遍历到33（即A33的行号）
        sheet1.row_dimensions[row_num].height = 20.25  # 设置行高

    # 设置字体
    font = Font(name='宋体', size=16, bold=True)
    sheet1.cell(row=bgs + KS + 1, column=1).font = font

    # 假设以下变量已经定义并赋值：型号, 出厂盘号, 出厂长度, 光缆芯数, 光缆长度
    # 填写表头
    header_rows = [
        ["附表2 光缆单盘检验测试记录表", "", "", "", "", ""],  # 第2行表头
        ["光缆型号", 型号, "出厂盘号", 出厂盘号, "出厂长度", 出厂长度],  # 第3行表头数据
        ["光缆芯数", str(光缆芯数) + "B1", "光缆长度", 光缆长度, "外观", "完整"],  # 第4行表头数据
        ["测试仪表", "AQ7260", "单位", "km", "≤0.22dB/km", ""],  # 第5行表头数据
        ["纤芯序号", "1550nm（折射率1.4680）", "", "纤芯序号", "1550nm（折射率1.4680）", ""],  # 第6行表头数据
        ["", "测试衰耗值", "测试光纤长度", "", "测试衰耗值", "测试光纤长度"],  # 第7行表头数据
    ]
    start_row = bgs + KS + 1
    for row_data in header_rows:
        for col_index, value in enumerate(row_data, start=1):
            sheet1.cell(row=start_row, column=col_index).value = value
        start_row += 1

    # 页尾
    footer_rows = [
        ["结论", "", "", "", "", ""],
        ["测试人员：", "", "", "监理人员：", "", ""],  # 第32行页尾
        ["日期： ", "", "", "日期： ", "", ""]  # 第33行页尾
    ]
    start_row1 = bgs + KS + 31
    for row_data in footer_rows:
        for col_index, value in enumerate(row_data, start=1):
            sheet1.cell(row=start_row1, column=col_index).value = value
        start_row1 += 1

        def 测试衰耗值():
            return round(random.uniform(0.185, 0.2), 3)

        # 光缆芯数 = 8
        if 光缆芯数 in [48, 96, 144]:
            for m in range(24):
                sheet1.cell(row=bgs + KS + m + 7, column=1).value = m + 1 + 96
                sheet1.cell(row=bgs + KS + m + 7, column=4).value = m + 25 + 96
                sheet1.cell(row=bgs + KS + m + 7, column=2).value = 测试衰耗值()
                sheet1.cell(row=bgs + KS + m + 7, column=5).value = 测试衰耗值()
                sheet1.cell(row=bgs + KS + m + 7, column=3).value = 测试光纤长度
                sheet1.cell(row=bgs + KS + m + 7, column=6).value = 测试光纤长度

        else:
            for i in [4, 8, 12, 24]:
                if 光缆芯数 == i:
                    for m in range(i):
                        sheet1.cell(row=bgs + KS + m + 7, column=1).value = m + 1
                        sheet1.cell(row=bgs + KS + m + 7, column=2).value = 测试衰耗值()
                        sheet1.cell(row=bgs + KS + m + 7, column=3).value = 测试光纤长度
                    break
            # else:
            # print("Invalid input")

    # 定义要合并的单元格区域
    merge_ranges = [
        "A{}:F{}".format(bgs + KS + 1, bgs + KS + 1),
        "B{}:C{}".format(bgs + KS + 5, bgs + KS + 5),
        "E{}:F{}".format(bgs + KS + 4, bgs + KS + 4),
        "E{}:F{}".format(bgs + KS + 5, bgs + KS + 5),
        "A{}:A{}".format(bgs + KS + 5, bgs + KS + 6),
        "D{}:D{}".format(bgs + KS + 5, bgs + KS + 6),
        "B{}:F{}".format(bgs + KS + 31, bgs + KS + 31)
    ]

    # 循环遍历合并单元格
    for merge_range in merge_ranges:
        sheet1.merge_cells(merge_range)

        # 设置 A1:F31 区域的文字上下左右居中
    alignment_center = Alignment(horizontal='center', vertical='center')
    for row in range(bgs + KS + 1, bgs + KS + 34):
        for col in range(1, 7):
            cell = sheet1.cell(row=row, column=col)
            cell.alignment = alignment_center

            # 设置单元格格式，保留3位小数
    number_format = ('0.000')

    sheet1.cell(bgs + KS + 2, 6).number_format = number_format
    sheet1.cell(bgs + KS + 3, 4).number_format = number_format

    # B7:B30 和 E7:E30 范围内的单元格设置格式
    for row in range(bgs + KS + 7, bgs + KS + 31):
        for col in [2, 3, 5, 6]:  # 列B和列E的列号
            cell = sheet1.cell(row=row, column=col)
            cell.number_format = number_format

            # 添加边框线
    border = Border(left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin'))
    for row in range(bgs + KS + 2, bgs + KS + 32):
        for col in range(1, 7):
            cell = sheet1.cell(row=row, column=col)
            cell.border = border


def my_function48(sheet1, 型号, 光缆芯数, 页数):
    N = 0  # 初始值
    for i in range(页数):  # 假设我们想要循环10次
        KS = N * 33

        # 获取当前日期
        now = datetime.now()
        six_months_ago = now - relativedelta(months=6)
        random_month_offset = random.randint(0, 5)
        target_month_days = []

        # 计算目标月份的最大天数
        for offset in range(random_month_offset + 1):
            target_month = six_months_ago + relativedelta(months=offset)
            target_month_days.append(calendar.monthrange(target_month.year, target_month.month)[1])

        random_day_offset = random.randint(1, min(now.day, *target_month_days))

        # 现在可以安全地设置随机日期
        random_date = six_months_ago + relativedelta(months=random_month_offset)
        random_date = random_date.replace(day=random_day_offset)

        random_date_str = random_date.strftime('%Y%m%d')
        random_number = random.randint(1, 999)
        sjsz = str(random_number).zfill(3)

        出厂盘号 = "C" + random_date_str + str(sjsz)

        # 填入对应数据
        # 生成一个2000至2020之间的随机数，数字取整
        def 长度():
            # return random.randint(2000, 2020)
            return round(random.uniform(2, 2.010), 3)

        测试光纤长度 = 长度()

        第一个表(sheet1, 型号, 光缆芯数, 出厂盘号, KS, 测试光纤长度)  # 调用第一个表

        N += 1  # 每次循环，N加1


def my_function96(sheet1, 型号, 光缆芯数, 页数):
    N = 0  # 初始值
    for i in range(页数):  # 假设我们想要循环10次
        KS = N * 33 * 2

        # 获取当前日期
        now = datetime.now()
        six_months_ago = now - relativedelta(months=6)
        random_month_offset = random.randint(0, 5)
        target_month_days = []

        # 计算目标月份的最大天数
        for offset in range(random_month_offset + 1):
            target_month = six_months_ago + relativedelta(months=offset)
            target_month_days.append(calendar.monthrange(target_month.year, target_month.month)[1])

        random_day_offset = random.randint(1, min(now.day, *target_month_days))

        # 现在可以安全地设置随机日期
        random_date = six_months_ago + relativedelta(months=random_month_offset)
        random_date = random_date.replace(day=random_day_offset)

        random_date_str = random_date.strftime('%Y%m%d')
        random_number = random.randint(1, 999)
        sjsz = str(random_number).zfill(3)

        出厂盘号 = "C" + random_date_str + str(sjsz)

        # 填入对应数据
        # 生成一个2000至2020之间的随机数，数字取整
        def 长度():
            # return random.randint(2000, 2020)
            return round(random.uniform(2, 2.010), 3)

        测试光纤长度 = 长度()

        第一个表(sheet1, 型号, 光缆芯数, 出厂盘号, KS, 测试光纤长度)  # 调用第一个表
        第二个表(sheet1, 型号, 光缆芯数, 出厂盘号, KS, 测试光纤长度)

        N += 1  # 每次循环，N加1


def my_function144(sheet1, 型号, 光缆芯数, 页数):
    N = 0  # 初始值
    for i in range(页数):  # 假设我们想要循环10次
        KS = N * 33 * 3

        # 获取当前日期
        now = datetime.now()
        six_months_ago = now - relativedelta(months=6)
        random_month_offset = random.randint(0, 5)
        target_month_days = []

        # 计算目标月份的最大天数
        for offset in range(random_month_offset + 1):
            target_month = six_months_ago + relativedelta(months=offset)
            target_month_days.append(calendar.monthrange(target_month.year, target_month.month)[1])

        random_day_offset = random.randint(1, min(now.day, *target_month_days))

        # 现在可以安全地设置随机日期
        random_date = six_months_ago + relativedelta(months=random_month_offset)
        random_date = random_date.replace(day=random_day_offset)

        random_date_str = random_date.strftime('%Y%m%d')
        random_number = random.randint(1, 999)
        sjsz = str(random_number).zfill(3)

        出厂盘号 = "C" + random_date_str + str(sjsz)

        # 填入对应数据
        # 生成一个2000至2020之间的随机数，数字取整
        def 长度():
            # return random.randint(2000, 2020)
            return round(random.uniform(2, 2.010), 3)

        测试光纤长度 = 长度()

        第一个表(sheet1, 型号, 光缆芯数, 出厂盘号, KS, 测试光纤长度)  # 调用第一个表
        第二个表(sheet1, 型号, 光缆芯数, 出厂盘号, KS, 测试光纤长度)
        第三个表(sheet1, 型号, 光缆芯数, 出厂盘号, KS, 测试光纤长度)

        N += 1  # 每次循环，N加1
