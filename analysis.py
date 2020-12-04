# 读取与计算优分率，差分率，合格率
import xlrd
import os
import pandas as pd
import collections
from config import class_replace
import xlwt
import datetime
from xlutils.copy import copy

all_subjects = ['语文', '数学', '英语', '物理', '化学', '政治', '历史', '地理', '生物']
subjects_manfen = {'语文': 100, '数学': 100, '英语': 100, '物理': 100, '化学': 80, '政治': 60, '历史': 60, '地理': 50, '生物': 50}
write_cols_order = ["姓名", "班级", "语文", "数学", "英语", "物理", '化学', "政治", "历史", '地理', '生物', "总分", "名次"]


def find_the_excel(dir_name="初二成绩"):
    """
    找到文件夹下面最可能是需要计算的excel文件
    :return:
    """

    if len(list(os.listdir(dir_name))) >= 1:
        for file in os.listdir(dir_name):
            if file.endswith("xls") or file.endswith("xlsx"):
                if "老师班级配置" not in file:
                    return os.path.join(dir_name, file)
    print("原始数据有误")
    exit(-1)


def read_teacher_name(dir_name="初二成绩"):
    peizhi_file = os.path.join(dir_name, "老师班级配置.xls")
    if os.path.exists(peizhi_file):
        df = pd.read_excel(peizhi_file, sheet_name="Sheet1", header=0, index_col=0)
        return df.to_dict()
    else:
        return None


def classify_with_class(excel_path, result_file="result.xls", dele=True):
    # workshop = xlrd.open_workbook(excel_path)
    # sheet1 = workshop.sheet_by_name("Sheet1")
    # print([i.value for i in sheet1.row(0)])
    df = pd.read_excel(excel_path, sheet_name="Sheet1", header=0)
    # df[["班级"]].replace(class_replace, inplace=True)
    df[["班级"]] = df[["班级"]].applymap(lambda x: class_replace[x] if x in class_replace else x)
    origin_len = len(df)
    columns = df.columns.values
    only_subjects = list(set(columns) & set(all_subjects))
    subjects = only_subjects + ["总分"]

    write_cols = list(sorted(set(columns) & set(write_cols_order), key=lambda x: write_cols_order.index(x)))

    tmp = df[only_subjects]
    # tmp.replace({"缺考": 0}, inplace=True)
    quekao = {"缺考": 0}
    tmp = tmp.applymap(lambda x: quekao[x] if x in quekao else x)
    df["总分"] = tmp.apply(lambda x: x.sum(), axis=1)
    df.sort_values(by='总分', axis=0, ascending=False, inplace=True, na_position='last')
    df["名次"] = range(1, len(df) + 1)

    for col_name in df:
        if col_name in subjects:
            if dele:
                df.drop(df[df[col_name] == "缺考"].index, inplace=True)
            else:
                df[col_name] = df[col_name].map(lambda x: 0 if x == "缺考" else x)

    filter_len = len(df)
    print("应考人数：{}".format(origin_len))
    print("实考人数：{}".format(filter_len))
    print("缺考人数：{}".format(origin_len - filter_len))

    bak = df[subjects].mean()

    averages = dict()
    jigelvs = dict()
    youfenlvs = dict()
    chafenlvs = dict()
    zongfen = sum(subjects_manfen[i] for i in only_subjects)
    subjects_manfen_tmp = subjects_manfen.copy()
    subjects_manfen_tmp.update({"总分": zongfen})
    for col_name in df:
        if col_name in subjects:
            averages[col_name] = bak.loc[col_name]
            jigelvs[col_name] = len(df[df[col_name] >= subjects_manfen_tmp[col_name] * 0.6]) / (
                len(df[col_name]))  # 减去均分那一行
            youfenlvs[col_name] = len(df[df[col_name] >= subjects_manfen_tmp[col_name] * 0.8]) / (len(df[col_name]))
            chafenlvs[col_name] = len(df[df[col_name] < subjects_manfen_tmp[col_name] * 0.4]) / (len(df[col_name]))

    class_info_dict = collections.OrderedDict()
    writer = pd.ExcelWriter(result_file, engine="openpyxl")
    df[subjects] = df[subjects].astype(float)
    tmp = df.groupby('班级')[subjects].mean()
    df.sort_values(by='总分', axis=0, ascending=False, inplace=True, na_position='last')
    sorted_average = dict()
    qian_yibailiu_fens = dict()
    hou_yibailiu_fens = dict()
    for col_name in subjects:
        sorted_average[col_name] = sorted(list(zip(tmp.index, tmp[col_name])), key=lambda x: x[1], reverse=True)
        tt = df.sort_values(by=col_name, axis=0, ascending=False, inplace=False, na_position='last')
        qian_yibailiu_fens[col_name] = dict(tt.iloc[:160].groupby("班级").count()[col_name])
        hou_yibailiu_fens[col_name] = dict(tt.iloc[-160:].groupby("班级").count()[col_name])
    for _, class_df in df.groupby(df['班级']):
        class_subjects = dict()
        isinstance(class_df, pd.DataFrame)
        bak_tmp = class_df[subjects].mean()
        for col_name in subjects:
            jigelv = len(class_df[class_df[col_name] >= subjects_manfen_tmp[col_name] * 0.6]) / (
                len(class_df[col_name]))
            youfenlv = len(class_df[class_df[col_name] >= subjects_manfen_tmp[col_name] * 0.8]) / (
                len(class_df[col_name]))
            chafenlv = len(class_df[class_df[col_name] < subjects_manfen_tmp[col_name] * 0.4]) / (
                len(class_df[col_name]))
            junfen = bak_tmp.loc[col_name]
            subject = {"均分": junfen, "差分率": chafenlv, "合格率": jigelv, "优分率": youfenlv,
                       '排名': sorted_average[col_name].index((_, tmp.loc[_, col_name])) + 1, "班级": _,
                       "前160": qian_yibailiu_fens[col_name][_],
                       "后160": hou_yibailiu_fens[col_name][_]
                       }
            class_subjects[col_name] = subject
        class_info_dict[_] = (class_subjects, len(class_df))
        for i in ["优分率", "合格率", "差分率", "均分"]:
            tmp_ = {"姓名": i}
            for co in subjects:
                tmp_[co] = class_subjects[co][i]
            s = pd.DataFrame(tmp_, index=[len(class_df) + 1])
            class_df = class_df.append(s)
        write_class_df = class_df[write_cols]
        write_class_df.to_excel(writer, encoding='utf-8', sheet_name='{}班'.format(_), index=None)
        writer.save()
        writer.close()
    return class_info_dict, averages, jigelvs, chafenlvs, youfenlvs


def write_to_grade_one_excel(class_info_dict, averages, hegelvs, chafenlvs, youfenlvs, result_file="result.xls",
                             dir_name="初二成绩"):
    teachers = read_teacher_name(dir_name)
    data = xlrd.open_workbook(result_file)
    workbook = copy(wb=data)  # 完成xlrd对象向xlwt对象转换
    sheet = workbook.add_sheet("数据分析表")
    cols = ["教师", "排名", "均分", "合格率", "优分率", "差分率", "前160", "后160"]
    borders = xlwt.Borders()
    borders.left = 1
    borders.right = 1
    borders.top = 1
    borders.bottom = 1
    borders.bottom_colour = 0x3A

    juzhongstyle = xlwt.Alignment()
    juzhongstyle.horz = 0x02  # 设置水平居中
    juzhongstyle.vert = 0x01  # 设置垂直居中

    front_class_title = xlwt.Font()
    front_class_title.name = "微软雅黑"
    front_class_title.bold = True
    front_class_title.height = 20 * 15

    class_style = xlwt.XFStyle()  # 创建一个样式对象，初始化样式
    class_style.alignment = juzhongstyle
    class_style.font = front_class_title
    class_style.borders = borders

    year = datetime.datetime.now().year
    month = datetime.datetime.now().month
    day = datetime.datetime.now().day

    xueqi = "一" if 3 > month or month > 8 else "二"
    year = year if 3 > month or month > 8 else year - 1
    title = "{}-{}学年度第{}学期初一年级期末调研质量分析   {}.{}月".format(year, year + 1, xueqi, year, month)

    blocks = ['语文', '数学', '英语', '政治']
    first_colmns = ['语 文 (100)', '数 学 (100)', '英 语 (100)', '政 治 (60)']
    sheet.write_merge(0, 0, 0, len(blocks) * len(cols) + 1, title, class_style)  # Merges row 0's columns 0 through 3
    begin_row = 2

    block_writer_utils(blocks, first_colmns, cols, sheet, begin_row, class_info_dict, averages, hegelvs, chafenlvs,
                       youfenlvs, teachers, class_prefix="七")
    begin_row = begin_row + len(class_info_dict) + 3
    blocks = ['历史', '地理', '地理', '生物', '总分']
    first_colmns = ['历 史 (60)', '地 理 (50)', '生 物 (50)', '总 分 (520)']
    block_writer_utils(blocks, first_colmns, cols, sheet, begin_row, class_info_dict, averages, hegelvs, chafenlvs,
                       youfenlvs, teachers, class_prefix="七")
    workbook.save(result_file)


def write_to_grade_two_excel(class_info_dict, averages, hegelvs, chafenlvs, youfenlvs, result_file="result.xls",
                             dir_name="初二成绩"):
    teachers = read_teacher_name(dir_name)
    data = xlrd.open_workbook(result_file)
    workbook = copy(wb=data)  # 完成xlrd对象向xlwt对象转换
    sheet = workbook.add_sheet("数据分析表")
    cols = ["教师", "排名", "均分", "合格率", "优分率", "差分率", "前160", "后160"]
    borders = xlwt.Borders()
    borders.left = 1
    borders.right = 1
    borders.top = 1
    borders.bottom = 1
    borders.bottom_colour = 0x3A

    juzhongstyle = xlwt.Alignment()
    juzhongstyle.horz = 0x02  # 设置水平居中
    juzhongstyle.vert = 0x01  # 设置垂直居中

    front_class_title = xlwt.Font()
    front_class_title.name = "微软雅黑"
    front_class_title.bold = True
    front_class_title.height = 20 * 15

    class_style = xlwt.XFStyle()  # 创建一个样式对象，初始化样式
    class_style.alignment = juzhongstyle
    class_style.font = front_class_title
    class_style.borders = borders

    year = datetime.datetime.now().year
    month = datetime.datetime.now().month
    day = datetime.datetime.now().day

    xueqi = "一" if 3 > month or month > 8 else "二"
    year = year if 3 > month or month > 8 else year - 1
    title = "{}-{}学年度第{}学期初二年级期末调研质量分析   {}.{}月".format(year, year + 1, xueqi, year, month)

    blocks = ['语文', '数学', '英语', '物理']
    first_colmns = ['语 文 (100)', '数 学 (100)', '英 语 (100)', '物 理 (100)']
    sheet.write_merge(0, 0, 0, len(blocks) * len(cols) + 1, title, class_style)  # Merges row 0's columns 0 through 3
    begin_row = 2

    block_writer_utils(blocks, first_colmns, cols, sheet, begin_row, class_info_dict, averages, hegelvs, chafenlvs,
                       youfenlvs, teachers, class_prefix="八")
    begin_row = begin_row + len(class_info_dict) + 3
    blocks = ['政治', '历史', '总分']
    first_colmns = ['政 治 (60)', '历 史 (60)', '总 分 (520)']
    block_writer_utils(blocks, first_colmns, cols, sheet, begin_row, class_info_dict, averages, hegelvs, chafenlvs,
                       youfenlvs, teachers, class_prefix="八")
    workbook.save(result_file)


def write_to_grade_three_excel(class_info_dict, averages, hegelvs, chafenlvs, youfenlvs, result_file="result.xls",
                               dir_name="初三成绩"):
    teachers = read_teacher_name(dir_name)
    data = xlrd.open_workbook(result_file)
    workbook = copy(wb=data)  # 完成xlrd对象向xlwt对象转换
    sheet = workbook.add_sheet("数据分析表")
    cols = ["教师", "排名", "均分", "合格率", "优分率", "差分率", "前160", "后160"]
    borders = xlwt.Borders()
    borders.left = 1
    borders.right = 1
    borders.top = 1
    borders.bottom = 1
    borders.bottom_colour = 0x3A

    juzhongstyle = xlwt.Alignment()
    juzhongstyle.horz = 0x02  # 设置水平居中
    juzhongstyle.vert = 0x01  # 设置垂直居中

    front_class_title = xlwt.Font()
    front_class_title.name = "微软雅黑"
    front_class_title.bold = True
    front_class_title.height = 20 * 15

    class_style = xlwt.XFStyle()  # 创建一个样式对象，初始化样式
    class_style.alignment = juzhongstyle
    class_style.font = front_class_title
    class_style.borders = borders

    year = datetime.datetime.now().year
    month = datetime.datetime.now().month
    day = datetime.datetime.now().day

    xueqi = "一" if 3 > month or month > 8 else "二"
    year = year if 3 > month or month > 8 else year - 1
    title = "{}-{}学年度第{}学期初三年级期末调研质量分析   {}.{}月".format(year, year + 1, xueqi, year, month)

    blocks = ['语文', '数学', '英语', '物理']
    first_colmns = ['语 文 (120)', '数 学 (120)', '英 语 (120)', '物 理 (100)']
    sheet.write_merge(0, 0, 0, len(blocks) * len(cols) + 1, title, class_style)  # Merges row 0's columns 0 through 3
    begin_row = 2

    block_writer_utils(blocks, first_colmns, cols, sheet, begin_row, class_info_dict, averages, hegelvs, chafenlvs,
                       youfenlvs, teachers, class_prefix="九")
    begin_row = begin_row + len(class_info_dict) + 3
    blocks = ['化学', '政治', '历史', '总分']
    first_colmns = ['化 学 (80)', '政 治 (60)', '历 史 (60)', '总 分 (660)']
    block_writer_utils(blocks, first_colmns, cols, sheet, begin_row, class_info_dict, averages, hegelvs, chafenlvs,
                       youfenlvs, teachers, class_prefix="九")
    workbook.save(result_file)


def block_writer_utils(blocks, first_colmns, cols, sheet, begin_row, class_info_dict, averages, hegelvs, chafenlvs,
                       youfenlvs, teachers=None, class_prefix="八"):
    borders = xlwt.Borders()
    borders.left = 1
    borders.right = 1
    borders.top = 1
    borders.bottom = 1
    borders.bottom_colour = 0x3A

    juzhongstyle = xlwt.Alignment()
    juzhongstyle.horz = 0x02  # 设置水平居中
    juzhongstyle.vert = 0x01  # 设置垂直居中

    front_class_title = xlwt.Font()
    front_class_title.name = "微软雅黑"
    front_class_title.bold = True
    front_class_title.height = 20 * 15

    front_cols_title = xlwt.Font()
    front_cols_title.name = "微软雅黑"
    front_cols_title.bold = True

    class_style = xlwt.XFStyle()  # 创建一个样式对象，初始化样式
    class_style.alignment = juzhongstyle
    class_style.font = front_class_title
    class_style.borders = borders

    cols_style = xlwt.XFStyle()  # 创建一个样式对象，初始化样式
    cols_style.alignment = juzhongstyle
    cols_style.font = front_cols_title
    cols_style.borders = borders

    baifenbi_style = xlwt.XFStyle()  # 创建一个样式对象，初始化样式
    baifenbi_style.alignment = juzhongstyle
    baifenbi_style.num_format_str = '0.00%'
    baifenbi_style.borders = borders

    two_remain_style = xlwt.XFStyle()  # 创建一个样式对象，初始化样式
    two_remain_style.alignment = juzhongstyle
    two_remain_style.num_format_str = '0.00'
    two_remain_style.borders = borders

    juzhong_style = xlwt.XFStyle()  # 创建一个样式对象，初始化样式
    juzhong_style.alignment = juzhongstyle
    juzhong_style.borders = borders

    sheet.write(begin_row, 0, "班级", cols_style)
    sheet.write(begin_row, 1, "人数", cols_style)
    for i in range(len(blocks)):
        for j in range(len(cols)):
            sheet.write(begin_row, i * len(cols) + 2 + j, cols[j], cols_style)
    sheet.write(begin_row - 1, 0, "学科", cols_style)

    for i, col_ in enumerate(first_colmns):
        sheet.write_merge(begin_row - 1, begin_row - 1, i * len(cols) + 2, (i + 1) * len(cols) + 1, col_, class_style)
    for class_index, class_name in enumerate(class_info_dict, start=1):
        row_num = begin_row + class_index
        sheet.write(row_num, 0, class_prefix + str(class_name), cols_style)
        sheet.write(row_num, 1, class_info_dict[class_name][1], juzhong_style)
        for block_index, block in enumerate(blocks, start=1):
            for col_index, col in enumerate(cols, start=1):
                col_num = 1 + len(cols) * (block_index - 1) + col_index
                if col in class_info_dict[class_name][0][block]:
                    if "率" in col:
                        sheet.write(row_num, col_num, float(class_info_dict[class_name][0][block][col]), baifenbi_style)
                    elif "分" in col:
                        sheet.write(row_num, col_num, float(class_info_dict[class_name][0][block][col]),
                                    two_remain_style)
                    else:
                        sheet.write(row_num, col_num, float(class_info_dict[class_name][0][block][col]), juzhong_style)
                if col == "教师" and teachers:
                    try:
                        tmp = block if block != "总分" else "班主任"
                        sheet.write(row_num, col_num, teachers[tmp][class_prefix + str(class_name)], juzhong_style)
                    except Exception as e:
                        continue

    junfen_row = begin_row + len(class_info_dict) + 1
    sheet.write(junfen_row, 0, "年均", cols_style)
    for block_index, block in enumerate(blocks, start=1):
        for col_index, col in enumerate(cols, start=1):
            col_num = 1 + len(cols) * (block_index - 1) + col_index
            if col == "均分":
                sheet.write(junfen_row, col_num, float(averages[block]), two_remain_style)
            if col == "合格率":
                sheet.write(junfen_row, col_num, float(hegelvs[block]), baifenbi_style)
            if col == "差分率":
                sheet.write(junfen_row, col_num, float(chafenlvs[block]), baifenbi_style)
            if col == "优分率":
                sheet.write(junfen_row, col_num, float(youfenlvs[block]), baifenbi_style)


def handle_chuer():
    print("正在处理初二的成绩")
    dir_name = "初二成绩"
    subjects_manfen.update({"语文": 100, "数学": 100, "英语": 100})
    excel_path = find_the_excel(dir_name)
    print(excel_path)
    result_file = "GradeTworesult.xls"
    class_info_dict, averages, hegelvs, chafenlvs, youfenlvs = classify_with_class(excel_path, result_file=result_file)
    write_to_grade_two_excel(class_info_dict, averages, hegelvs, chafenlvs, youfenlvs, result_file=result_file,
                             dir_name=dir_name)


def handle_chuyi():
    print("正在处理初一的成绩")
    subjects_manfen.update({"语文": 100, "数学": 100, "英语": 100})
    dir_name = "初一成绩"
    excel_path = find_the_excel(dir_name)
    print(excel_path)
    result_file = "GradeOneresult.xls"
    class_info_dict, averages, hegelvs, chafenlvs, youfenlvs = classify_with_class(excel_path, result_file=result_file)
    write_to_grade_one_excel(class_info_dict, averages, hegelvs, chafenlvs, youfenlvs, result_file=result_file,
                             dir_name=dir_name)


def handle_chusan():
    print("正在处理初三的成绩")
    dir_name = "初三成绩"
    subjects_manfen.update({"语文": 120, "数学": 120, "英语": 120})
    excel_path = find_the_excel(dir_name)
    print(excel_path)
    result_file = "GradeThreeresult.xls"
    class_info_dict, averages, hegelvs, chafenlvs, youfenlvs = classify_with_class(excel_path, result_file=result_file)
    write_to_grade_three_excel(class_info_dict, averages, hegelvs, chafenlvs, youfenlvs, result_file=result_file,
                               dir_name=dir_name)


if __name__ == '__main__':
    print('\n'.join([''.join([('Love'[(x - y) % len('Love')]
                               if ((x * 0.05) ** 2 + (y * 0.1) ** 2 - 1) ** 3 - (x * 0.05) ** 2 * (
            y * 0.1) ** 3 <= 0 else ' ')  # 此处是根据心形曲线公式来的(x2+y2-1)3-x2y3=0
                              for x in range(-30, 30)])  # 定义高
                     for y in range(30, -30, -1)]))  # 定义宽
    print("这是给鸣夏小朋友的礼物")
    print()

    if os.path.exists("初一成绩"):
        handle_chuyi()

    if os.path.exists("初二成绩"):
        handle_chuer()

    if os.path.exists("初三成绩"):
        handle_chusan()

    # import time
    #
    # wait_time = 5
    # print("程序将于{}秒后退出".format(wait_time))
    # time.sleep(wait_time)
    import msvcrt

    print("按任意键退出。。。。")

    while True:
        if msvcrt.getch():
            break
