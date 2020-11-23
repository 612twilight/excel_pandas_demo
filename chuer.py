# 读取与计算优分率，差分率，合格率
import xlrd
import os
import pandas as pd
import collections
from config import class_replace
import xlwt

only_subjects = ['语文', '数学', '英语', '物理', '政治', '历史']
subjects = ['语文', '数学', '英语', '物理', '政治', '历史', '总分']
subjects_manfen = {'语文': 100, '数学': 100, '英语': 100, '物理': 100, '政治': 60, '历史': 60, '总分': 520}
write_cols = ["姓名", "班级", "语文", "数学", "英语", "物理", "政治", "历史", "总分", "名次"]
calculation_mean = ['语文', '数学', '英语', '物理', '政治', '历史', "总分"]


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


def classify_with_class(excel_path, dele=True):
    # workshop = xlrd.open_workbook(excel_path)
    # sheet1 = workshop.sheet_by_name("Sheet1")
    # print([i.value for i in sheet1.row(0)])
    df = pd.read_excel(excel_path, sheet_name="Sheet1", header=0)
    df["班级"].replace(class_replace, inplace=True)
    origin_len = len(df)
    tmp = df[only_subjects]
    # tmp.iloc[tmp == "缺考"] = 0
    df["总分"] = tmp.apply(lambda x: x.sum(), axis=1)
    df.sort_values(by='总分', axis=0, ascending=False, inplace=True, na_position='last')
    df["名次"] = range(1, len(df) + 1)

    # for col_name in df:
    #     if col_name in subjects:
    #         if dele:
    #             df.drop(df[df[col_name] == "缺考"].index, inplace=True)
    #         else:
    #             df[col_name] = df[col_name].map(lambda x: 0 if x == "缺考" else x)

    filter_len = len(df)
    print("应考人数：{}".format(origin_len))
    print("实考人数：{}".format(filter_len))
    print("缺考人数：{}".format(origin_len - filter_len))

    bak = df[calculation_mean].mean()

    averages = dict()
    jigelvs = dict()
    youfenlvs = dict()
    chafenlvs = dict()
    for col_name in df:
        if col_name in subjects:
            averages[col_name] = bak.loc[col_name]
            jigelvs[col_name] = len(df[df[col_name] >= subjects_manfen[col_name] * 0.6]) / (
                len(df[col_name]))  # 减去均分那一行
            youfenlvs[col_name] = len(df[df[col_name] >= subjects_manfen[col_name] * 0.8]) / (len(df[col_name]))
            chafenlvs[col_name] = len(df[df[col_name] < subjects_manfen[col_name] * 0.4]) / (len(df[col_name]))

    class_info_dict = collections.OrderedDict()
    writer = pd.ExcelWriter("result.xls", engine="openpyxl")
    df[calculation_mean] = df[calculation_mean].astype(float)
    tmp = df.groupby('班级')[calculation_mean].mean()
    df.sort_values(by='总分', axis=0, ascending=False, inplace=True, na_position='last')
    sorted_average = dict()
    qian_yibailiu_fens = dict()
    hou_yibailiu_fens = dict()
    for col_name in subjects:
        sorted_average[col_name] = sorted(list(zip(tmp.index, tmp[col_name])), key=lambda x: x[1], reverse=True)
        tt = df.sort_values(by=col_name, axis=0, ascending=False, inplace=False, na_position='last')
        qian_yibailiu_fens[col_name] = dict(tt.iloc[:160].groupby("班级").count()[col_name])
        hou_yibailiu_fens[col_name] = dict(tt.iloc[-160:].groupby("班级").count()[col_name])
    print(qian_yibailiu_fens)
    for _, class_df in df.groupby(df['班级']):
        class_subjects = dict()
        isinstance(class_df, pd.DataFrame)
        bak_tmp = class_df[calculation_mean].mean()
        # print(class_df.tail)

        for col_name in subjects:
            # ss = sorted(list(zip(tmp.index, tmp[col_name])), key=lambda x: x[1], reverse=True)
            if col_name == "政治" and _ == "十六班":
                print(len(class_df[class_df[col_name] >= subjects_manfen[col_name] * 0.6]))
            jigelv = len(class_df[class_df[col_name] >= subjects_manfen[col_name] * 0.6]) / (
                len(class_df[col_name]))
            youfenlv = len(class_df[class_df[col_name] >= subjects_manfen[col_name] * 0.8]) / (
                len(class_df[col_name]))
            chafenlv = len(class_df[class_df[col_name] < subjects_manfen[col_name] * 0.4]) / (
                len(class_df[col_name]))
            junfen = bak_tmp.loc[col_name]
            subject = {"均分": junfen, "差分率": chafenlv, "合格率": jigelv, "优分率": youfenlv,
                       '排名': sorted_average[col_name].index((_, tmp.loc[_, col_name])) + 1, "班级": _,
                       "前160": qian_yibailiu_fens[col_name][_],
                       "后160": hou_yibailiu_fens[col_name][_]}
            class_subjects[col_name] = subject
            class_info_dict[_] = (class_subjects, len(class_df))
            write_class_df = class_df[write_cols].copy()
            write_class_df.to_excel(writer, encoding='utf-8', sheet_name='{}班'.format(_), index=None)
            writer.save()
            writer.close()
            # print(df)
    return class_info_dict, averages, jigelvs, chafenlvs, youfenlvs


def write_to_another_excel(class_info_dict, averages, hegelvs, chafenlvs, youfenlvs):
    import datetime
    teachers = read_teacher_name()
    # workbook = xlwt.Workbook()
    from xlutils.copy import copy

    data = xlrd.open_workbook("result.xls")

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
    xueqi = "一" if month <= 9 else "二"
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
    workbook.save("result.xls")


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
                    except:
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


if __name__ == '__main__':
    excel_path = find_the_excel()
    print(excel_path)
    class_info_dict, averages, hegelvs, chafenlvs, youfenlvs = classify_with_class(excel_path)
    write_to_another_excel(class_info_dict, averages, hegelvs, chafenlvs, youfenlvs)
    # read_teacher_name()
