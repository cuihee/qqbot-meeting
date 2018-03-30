import datetime
from openpyxl import load_workbook, Workbook
import os
import re


def my_watch_group(contact, group_name):
    return contact.nick == group_name

# todo 考虑使用redis json存储规则
def dialog_clearify(content):
    """
    这里面的顺序不要乱
    :param content:
    :return:
    """
    d = content
    if not isinstance(d, type("字符串类型")):
        return ""
    clearify_dict = {'~~~': '-', "~~": "-", '  ': ' ',
                     '~~': '-', '~': '-', '---': '-',
                     '--': '-', '——': '-', '：': ':',
                     '全天': '8:30-18:00',
                     '12层': '12楼', '13层': '13楼',
                     '今日': '今天', '明日': '明天',
                     '合昌': '和昌', '合唱': '和昌', '和唱': '和昌',
                     '点半': ':30', '点30': ':30', '点三十': ':30',
                     '点': ':00',
                     '十一': '11', '十二': '12', '十三': '13',
                     '十四': '14', '十五': '15', '十六': '16',
                     '十七': '17', '十八': '18', '十九': '19',
                     '二十': '20', '十': '10',
                     '九': '9', '八': '8', '七': '7',
                     '六': '6', '五': '5', '四': '4',
                     '三': '3', '二': '2', '一': '1',
                     '下午1:': '13:', '下午2:': '14:', '下午3:': '15:', '下午4:': '16:',
                     '下午5:': '17:', '下午6:': '18:', '下午7:': '19:', '下午8:': '20:',
                     '预订': '预定',
                     '到': '-',
                     '（': '(', '）': ')',
                     '，': '.', ',': '.', '。': '.'
                     }
    for (k, v) in clearify_dict.items():
        d = d.replace(k, v)
    return d


def is_cmd(dialog):
    if not isinstance(dialog, type("")):
        return ""
    if dialog.find("预定") > -1 and dialog.find("会议室") > -1:
        return dialog


def find_fangjian(dialog):
    """
    用正则表达式从dialog中取出定义好的会议室名称
    http://www.runoob.com/regexp/regexp-metachar.html
    正则表达式教程↑

    :param dialog: 传入的字符串
    :return: 传出会议室名字
    """
    meeting_room = '12楼大会议室'
    r1 = re.findall(r'([1][23]楼)', dialog)
    r2 = re.findall(r'([大小]?会议室)', dialog)
    if r1.__len__() > 0:
        meeting_room = r1[0] + meeting_room[3:]
    if r2.__len__() > 0:
        meeting_room = meeting_room[:3] + r2[0]
    return meeting_room


def find_shijian(dialog):
    st = et = None
    r1 = re.findall(r'([0-9]?[0-9])[:点：]([0-5][0-9])', dialog)
    if r1.__len__() >= 1:  # 至少有一个时间
        st = r1[0][0] + ':' + r1[0][1]
    if r1.__len__() >= 2:  # 有两个时间
        et = r1[1][0] + ':' + r1[1][1]
    return st, et


def find_yuding(dialog):
    return not (-1 < dialog.find("取消") < 6)


def find_riqi(dialog):
    findre = re.findall(r'([12]?[0-9]{1,2}[月.][123]?[0-9][日号])', dialog)  # 12月02日
    if len(findre) > 0:
        findre2 = findre[0]
        if findre2.find("月") == 1 or findre2.find(".") == 1:
            findre2 = "0" + findre2
        findre2 = findre2.replace("月", "-")
        findre2 = findre2.replace(".", "-")
        if findre2.find("日") == 4 or findre2.find("号") == 4:
            findre2 = findre2[:3] + "0" + findre2[3:]
        findre2 = findre2.replace("日", "")  # 12-02
        findre2 = findre2.replace("号", "")  # 12-02
        findre2 = datetime.date.today().__str__()[:5] + findre2  # 2018-12-02
        return findre2
    findre = re.findall(r'([123]?[0-9][日号])', dialog)  # 2号
    if len(findre) > 0:
        findre2 = findre[0]
        if findre2.find("日") == 1:
            findre2 = "0" + findre2  # 02号
        findre2 = findre2.replace("日", "")
        findre2 = findre2.replace("号", "")  # 02
        findre2 = datetime.date.today().__str__()[:8] + findre2  # 2018-03-02
        return findre2
    if -1 < dialog.find("今"):
        return datetime.date.today().__str__()
    elif -1 < dialog.find("明"):
        return (datetime.date.today() + datetime.timedelta(days=1)).__str__()
    elif -1 < dialog.find("后"):
        return (datetime.date.today() + datetime.timedelta(days=2)).__str__()
    return datetime.date.today()


def create_sheet(sheetname, file):
    """
    根据sheetname创建sheet
    创建第一行
    :param sheetname:
    :return:
    """
    if sheetname in file.get_sheet_names():
        return
    file.create_sheet(sheetname)
    sheet = file.get_sheet_by_name(sheetname)
    sheet.cell(row=1, column=1).value = datetime.date.today().__str__()  # [:8]+"01"
    writetime(sheet=sheet, startrow=2)
    ccolumn = 1
    meeting_roomnames = ["12楼小会议室", "12楼大会议室", "13楼会议室"]
    for i in meeting_roomnames:
        ccolumn = ccolumn + 1
        sheet.cell(row=1, column=ccolumn).value = i
    return sheet


def get_excel_row(sheet, today):
    """
    :param sheet:
    :param today: datetime.date.today().__str__()
    :return: row of today
    """
    find_ornot = False
    find_row = 0
    for i in range(1, sheet.max_row+1):
        if sheet.cell(row=i, column=1).value == today:
            find_ornot = True
            find_row = i
            break
    if find_ornot:
        writetime(sheet=sheet, startrow=find_row+1)
        return find_row
    else:
        find_row = sheet.max_row + 1
        sheet.cell(row=find_row, column=1).value = datetime.date.today().__str__()
        writetime(sheet=sheet, startrow=find_row + 1)
        return find_row


def writetime(sheet, startrow):
    """
    不包括today信息的一天时间
    :param sheet:
    :param startrow:
    :return:
    """
    # excel_sheet.cell(row=startrow, excel_column=1).value = datetime.date.today().__str__()  # [:8]+"01"
    m = ["00", "15", "30", "45"]
    h = [i.__str__() for i in range(8, 21, 1)]
    crow = startrow
    for i in h:
        for j in m:
            sheet.cell(row=crow, column=1).value = i + ":" + j + ":00"
            crow = crow + 1


def occupy_it(sheet, st, en, co, info="占用人信息"):
    for i in range(st, en, 1):
        _ = sheet.cell(column=co, row=i, value=info)


def get_dtime(st, et):
    """
    计算所给时间段距离8:00的格子数 默认15min
    请注意8:00是第一个格子
    :param st:
    :param et:
    :return:
    """
    a = int(st[:st.find(":")])
    b = int(st[st.find(":") + 1:])
    c = int(et[:et.find(":")])
    d = int(et[et.find(":") + 1:])
    ds = (a - 8) * 4 + b // 15 + 1
    de = (c - 8) * 4 + d // 15 + 1
    return ds, de


def get_excel_column(mr):
    a = -1
    if mr.find("12楼小会议室") > -1:
        a = 2
    elif mr.find("12楼大会议室") > -1:
        a = 3
    elif mr.find("13楼会议室") > -1:
        a = 4
    return a


def get_excel_file(filename):
    """
    :param filename:
    :return:
    """
    dir_files = os.listdir(os.getcwd())
    if filename in dir_files:
        wb = load_workbook(filename)
    else:
        wb = Workbook()
        wb.save(filename)
        wb = load_workbook(filename)
    return wb


def get_excel_sheet(riqi, file):
    sheetnames = file.get_sheet_names()  # 所有表名
    month_name = riqi[:7]
    if month_name in sheetnames:  # 存在
        sheet = file.get_sheet_by_name(name=month_name)
    else:
        sheet = create_sheet(month_name, file)
    return sheet


def is_occupied(sheet, start, end, column):
    busy = False  # 假设没占用
    busy_info = ""
    for i in range(start, end, 1):
        if sheet.cell(row=i, column=column).value is not None:
            busy_info = sheet.cell(row=i, column=column).value
            busy = True
            break
    return busy, busy_info


def deal_book(sheet, start, end, column, info, book, bot, contact):
    if book:
        occupied, occupied_info = is_occupied(sheet, start, end, column)
        if occupied:
            bot.SendTo(contact, "预定失败\"" + occupied_info + "\"")
            print("已被\"" + occupied_info + "\"占用")

        else:
            occupy_it(sheet, start, end, column, info)
            print("成功预定")
    else:  # todo 取消预定
        pass


def excel_file_close(file, name):
    file.save(filename=name)

