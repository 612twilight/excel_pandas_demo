# 读取与计算优分率，差分率，合格率
import xlrd
import os
import pandas as pd
import openpyxl

only_subjects = ['语文', '数学', '英语', '政治', '历史', '地理', '生物']
subjects = ['语文', '数学', '英语', '政治', '历史', '地理', '生物', '总分']
subjects_manfen = {'语文': 100, '数学': 100, '英语': 100, '政治': 60, '历史': 60, '地理': 50, '生物': 50, '总分': 520}

write_cols = ["序号", "班级", "姓名", "语文", "数学", "英语", "政治", "历史", "地理", "生物", "总分", "名次"]

calculation_mean = ["语文", "数学", "英语", "政治", "历史", "地理", "生物", "总分"]
import collections


def find_the_excel(dir_name="初一成绩"):
    """
    找到文件夹下面最可能是需要计算的excel文件
    :return:
    """
    if len(list(os.listdir(dir_name))) == 1:
        for file in os.listdir("初一成绩"):
            if file.endswith("xls") or file.endswith("xlsx"):
                return os.path.join(dir_name, file)
    print("没有找到原始数据")
    exit(-1)


def classify_with_class(excel_path, dele=True):
    workshop = xlrd.open_workbook(excel_path)
    sheet1 = workshop.sheet_by_name("Sheet1")
    print([i.value for i in sheet1.row(0)])
    df = pd.read_excel(excel_path, header=0)

    origin_len = len(df)
    tmp = df[only_subjects]
    tmp.iloc[tmp == "缺考"] = 0
    df["总分"] = tmp.apply(lambda x: x.sum(), axis=1)
    df.sort_values(by='总分', axis=0, ascending=False, inplace=True, na_position='last')
    df["名次"] = range(1, len(df) + 1)

    for col_name in df:
        if col_name in subjects:
            if dele:
                df.drop(df[df[col_name] == "缺考"].index, inplace=True)
            else:
                df[col_name] = df[col_name].map(lambda x: 0 if x == "缺考" else x)
    df.loc["均分"] = df[calculation_mean].mean()
    averages = dict()
    jigelvs = dict()
    youfenlvs = dict()
    chafenlvs = dict()
    for col_name in df:
        if col_name in subjects:
            averages[col_name] = df.loc["均分", col_name]
            jigelvs[col_name] = len(df[df[col_name] >= subjects_manfen[col_name] * 0.6]) / len(df[col_name])
            youfenlvs[col_name] = len(df[df[col_name] >= subjects_manfen[col_name] * 0.8]) / len(df[col_name])
            chafenlvs[col_name] = len(df[df[col_name] < subjects_manfen[col_name] * 0.4]) / len(df[col_name])
    filter_len = len(df)
    print("应考人数：{}".format(origin_len))
    print("实考人数：{}".format(filter_len))
    print("缺考人数：{}".format(origin_len - filter_len))
    class_info_dict = collections.OrderedDict()
    writer = pd.ExcelWriter("tmp.xls", engine="openpyxl")
    df[calculation_mean] = df[calculation_mean].astype(float)
    tmp = df.groupby('班级')[calculation_mean].mean()
    df.sort_values(by='总分', axis=0, ascending=False, inplace=True, na_position='last')
    sorted_average = dict()
    qian_yibailiu_fens = dict()
    hou_yibailiu_fens = dict()
    for col_name in subjects:
        sorted_average[col_name] = sorted(list(zip(tmp.index, tmp[col_name])), key=lambda x: x[1], reverse=True)
        qian_yibailiu_fens[col_name] = dict(df.iloc[:160].groupby("班级").count()["总分"])
        hou_yibailiu_fens[col_name] = dict(df.iloc[-160:].groupby("班级").count()["总分"])
    for _, class_df in df.groupby(df['班级']):
        class_subjects = dict()
        isinstance(class_df, pd.DataFrame)
        class_df.loc["均分"] = class_df[calculation_mean].mean()
        for col_name in subjects:
            # ss = sorted(list(zip(tmp.index, tmp[col_name])), key=lambda x: x[1], reverse=True)

            jigelv = len(class_df[class_df[col_name] >= subjects_manfen[col_name] * 0.6]) / len(class_df[col_name])
            youfenlv = len(class_df[class_df[col_name] >= subjects_manfen[col_name] * 0.8]) / len(class_df[col_name])
            chafenlv = len(class_df[class_df[col_name] < subjects_manfen[col_name] * 0.4]) / len(class_df[col_name])
            junfen = class_df.loc['均分', col_name]
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


import xlwt


def write_to_another_excel(class_info_dict, averages, hegelvs, chafenlvs, youfenlvs):
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet("数据分析表")
    cols = ["教师", "排名", "均分", "合格率", "优分率", "差分率", "前160", "后160"]
    title = "学情检测"
    sheet.write_merge(0, 0, 0, 21, title)  # Merges row 0's columns 0 through 3
    sheet.write(1, 0, "学科")
    begin_row = 2
    sheet.write(begin_row, 0, "班级")
    sheet.write(begin_row, 1, "人数")
    for i in range(4):
        for j in range(len(cols)):
            sheet.write(begin_row, i * len(cols) + 2 + j, cols[j])
    blocks = ['语文', '数学', '英语', '政治']
    first_colmns = ['语文100', '数学100', '英语100', '政治60']

    for i, col_ in enumerate(first_colmns):
        sheet.write_merge(begin_row - 1, begin_row - 1, i * len(cols) + 2, (i + 1) * len(cols) + 1, col_)

    for class_index, class_name in enumerate(class_info_dict, start=1):
        row_num = begin_row + class_index
        sheet.write(row_num, 0, class_name)
        sheet.write(row_num, 1, class_info_dict[class_name][1])
        for block_index, block in enumerate(blocks, start=1):
            for col_index, col in enumerate(cols, start=1):
                col_num = 1 + len(cols) * (block_index - 1) + col_index
                if col in class_info_dict[class_name][0][block]:
                    sheet.write(row_num, col_num, float(class_info_dict[class_name][0][block][col]))
    junfen_row = begin_row + len(class_info_dict) + 1
    sheet.write(junfen_row, 0, "年均")
    for block_index, block in enumerate(blocks, start=1):
        for col_index, col in enumerate(cols, start=1):
            col_num = 1 + len(cols) * (block_index - 1) + col_index
            if col == "均分":
                sheet.write(junfen_row, col_num, float(averages[block]))
            if col == "合格率":
                sheet.write(junfen_row, col_num, float(hegelvs[block]))
            if col == "差分率":
                sheet.write(junfen_row, col_num, float(chafenlvs[block]))
            if col == "优分率":
                sheet.write(junfen_row, col_num, float(youfenlvs[block]))

    begin_row = begin_row + len(class_info_dict) + 3
    sheet.write(begin_row, 0, "班级")
    sheet.write(begin_row, 1, "人数")
    for i in range(4):
        for j in range(len(cols)):
            sheet.write(begin_row, i * len(cols) + 2 + j, cols[j])
    blocks = ['历史', '地理', '生物', '总分']
    first_colmns = ['历史60', '地理50', '生物50', '总分520']

    for i, col_ in enumerate(first_colmns):
        sheet.write_merge(begin_row - 1, begin_row - 1, i * len(cols) + 2, (i + 1) * len(cols) + 1, col_)

    for class_index, class_name in enumerate(class_info_dict, start=1):
        row_num = begin_row + class_index
        sheet.write(row_num, 0, class_name)
        sheet.write(row_num, 1, class_info_dict[class_name][1])
        for block_index, block in enumerate(blocks, start=1):
            for col_index, col in enumerate(cols, start=1):
                col_num = 1 + len(cols) * (block_index - 1) + col_index
                if col in class_info_dict[class_name][0][block]:
                    sheet.write(row_num, col_num, float(class_info_dict[class_name][0][block][col]))
    junfen_row = begin_row + len(class_info_dict) + 1
    sheet.write(junfen_row, 0, "年均")
    for block_index, block in enumerate(blocks, start=1):
        for col_index, col in enumerate(cols, start=1):
            col_num = 1 + len(cols) * (block_index - 1) + col_index
            if col == "均分":
                sheet.write(junfen_row, col_num, float(averages[block]))
            if col == "合格率":
                sheet.write(junfen_row, col_num, float(hegelvs[block]))
            if col == "差分率":
                sheet.write(junfen_row, col_num, float(chafenlvs[block]))
            if col == "优分率":
                sheet.write(junfen_row, col_num, float(youfenlvs[block]))
    workbook.save("tt.xls")
    pass


if __name__ == '__main__':
    excel_path = find_the_excel()
    print(excel_path)
    class_info_dict, averages, jigelvs, chafenlvs, youfenlvs = classify_with_class(excel_path)
    write_to_another_excel(class_info_dict, averages, jigelvs, chafenlvs, youfenlvs)
