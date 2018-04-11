from meeting_tools import *


watch_group_name = ["0.0", '张志琳、刘文、李豪、']
excel_file_name = "testMeeting.xlsx"
data_file_name = 'data_file.xlsx'
cmd_list = ['效率助手下载地址', '查询今天会议室预订情况', '查询明天会议室预订情况', '/?',  # 0 1 2 3
            '顺丰的联系方式']


def onQQMessage(bot, contact, member, content):
    if not my_watch_group(contact=contact, group_name=watch_group_name):
        return
    if '机器人回复' in content:
        return
    if '[@ME]' in content:
        bot.SendTo(contact, " at命令测试中, 也可以用来测试插件是否启用了")
        if cmd_list[0] in content:
            bot.SendTo(contact, " eepm.sippr.cn ")
        if cmd_list[1] in content:
            report_info = ask_info(data_file_name, (datetime.date.today() + datetime.timedelta(days=0)).__str__())
            bot.SendTo(contact, report_info.__str__())
        if cmd_list[2] in content:
            report_info = ask_info(data_file_name, (datetime.date.today()+datetime.timedelta(days=1)).__str__())
            bot.SendTo(contact, report_info.__str__())
        if cmd_list[3] in content:
            bot.SendTo(contact, '下列引号内的命令采用精确匹配。' + cmd_list.__str__())
        if cmd_list[4] in content:
            bot.SendTo(contact, '顺丰 15936240735')

    dialog = dialog_clearify(content)
    dialog = is_cmd(dialog)
    if len(dialog) < 1:
        print("不是预定会议室")
        return
    else:
        yuding_info = member.name + ' 群"' + contact.nick + '" ' + datetime.datetime.today().__str__()[:-7] + '     '
        book_ornot = find_yuding(dialog)  # True False
        riqi = find_riqi(dialog)  # 2018-03-13
        # print("获取日期:", riqi)
        yuding_info = yuding_info + riqi + ' '

        start_time, end_time = find_shijian(dialog)  # 8:00, 9:15
        # print("获取时间:", start_time, end_time)
        yuding_info = yuding_info + start_time + '-' + end_time + ' '

        fangjian = find_fangjian(dialog)  # 和昌12楼小会议室
        # print("获取房间名:", get_meetingrooms_names()[fangjian])
        yuding_info = yuding_info + fangjian.__str__() + ' ' + get_meetingrooms_names()[fangjian] + ' '
        print(yuding_info)

        # 表格文件对象
        excel_file = get_excel_file(filename=excel_file_name)
        # print("获取文件:", excel_file)

        excel_sheet = get_excel_sheet(riqi=riqi, file=excel_file)
        # print("获取sheet:", excel_sheet)

        column0 = 2  # 第一列是时间 从第二列开始是会议室编号 12层小会议室0 12层大会议室1 13层会议室2
        excel_column = fangjian + column0
        # print("获取列:", excel_column)

        excel_date_row = get_excel_row(sheet=excel_sheet, today=riqi)  # 下一行是8:00:00
        # print("获取行:", excel_date_row)

        delta_row_start, delta_row_end = get_dtime(start_time, end_time)
        # print("获取行区段:", excel_date_row+delta_row_start, excel_date_row+delta_row_end)

        deal_book(sheet=excel_sheet, start=excel_date_row + delta_row_start,
                  end=excel_date_row + delta_row_end, column=excel_column,
                  info=yuding_info, book=book_ornot, bot=bot, contact=contact)

        # 存储表格文件
        excel_file.save(filename=excel_file_name)
        print("\n")
