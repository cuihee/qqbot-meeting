from meeting_tools import *


watch_group_name = ["0.0", '张志琳、刘文、李豪、']
excel_file_name = "testMeeting.xlsx"


def onQQMessage(bot, contact, member, content):
    if not my_watch_group(contact=contact, group_name=watch_group_name):
        return
    if '[@ME]' in content:
        bot.SendTo(contact, " at命令未启用, 但是可以用来测试插件是否启用了")

    dialog = dialog_clearify(content)
    dialog = is_cmd(dialog)
    if len(dialog) < 1:
        print("不是预定会议室")
    else:
        yuding_info = member.name
        book_ornot = find_yuding(dialog)  # True False
        riqi = find_riqi(dialog)  # 2018-03-13
        print("获取日期:", riqi)

        start_time, end_time = find_shijian(dialog)  # 8:00, 9:15
        print("获取时间:", start_time, end_time)

        fangjian = find_fangjian(dialog)  # 和昌12楼小会议室
        print("获取房间名:", get_meetingrooms_names()[fangjian])

        # 表格文件对象
        excel_file = get_excel_file(filename=excel_file_name)
        print("获取文件:", excel_file)

        excel_sheet = get_excel_sheet(riqi=riqi, file=excel_file)
        print("获取sheet:", excel_sheet)

        column0 = 2  # 第一列是时间 从第二列开始是会议室编号 12层小会议室0 12层大会议室1 13层会议室2
        excel_column = fangjian + column0
        print("获取列:", excel_column)

        excel_date_row = get_excel_row(sheet=excel_sheet, today=riqi)  # 下一行是8:00:00
        print("获取行:", excel_date_row)

        delta_row_start, delta_row_end = get_dtime(start_time, end_time)
        print("获取行区段:", excel_date_row+delta_row_start, excel_date_row+delta_row_end)

        deal_book(sheet=excel_sheet, start=excel_date_row + delta_row_start,
                  end=excel_date_row + delta_row_end, column=excel_column,
                  info=yuding_info, book=book_ornot, bot=bot, contact=contact)

        # 存储表格文件
        excel_file.save(filename=excel_file_name)
        print("\n")
