def tr_digit_to_zn(digit):
    # 940,2400,0452
    digit = str(digit)
    length = len(digit)
    digit = digit[::-1]
    split = []
    sp_nums = range(0, length, 4)
    for i in sp_nums:
        split.append(digit[i: i + 4][::-1].zfill(4))
    # print(split)
    d_digit_to_zn = {
        0: "零",
        1: "一",
        2: "二",
        3: "三",
        4: "四",
        5: "五",
        6: "六",
        7: "七",
        8: "八",
        9: "九",
    }
    res_zn_list = []
    split_count = len(split)
    for i, e in enumerate(split):
        zn = ''
        for j, each in enumerate(e):
            if each == '0':
                if j == 0 and i == split_count - 1:
                    pass
                elif e[j - 1] == '0':
                    pass
                elif e[j:].strip('0'):
                    zn += '零'
            else:
                zn += d_digit_to_zn[int(each)] + {0: '千', 1: '百', 2: '十', 3: ''}[j]
        zn = zn + {0: '', 1: '万', 2: '亿'}[i]
        res_zn_list.append(zn)
    res_zn_list.reverse()
    res_zn = ''.join(res_zn_list)
    # print(res_zn)

    res_zn = [e for e in res_zn]
    for i, e in enumerate(res_zn):
        if e in '百千':
            try:
                if res_zn[i - 1] == '二':
                    res_zn[i - 1] = '两'
            except:
                pass
    res_zn = ''.join(res_zn)

    if res_zn.startswith('一十'):
        res_zn = res_zn[1:]

    if res_zn.startswith('二') and len(res_zn) >= 2 and res_zn[1] in ['万', '亿']:
        res_zn = '两' + res_zn[1:]

    return res_zn


class_replace = {tr_digit_to_zn(i) + "班": i for i in range(1, 100)}
