from .meeting_tools import *


watch_group_name = "0.0"
excel_file_name = "testMeeting.xlsx"



def onQQMessage(bot, contact, member, content):
    if not my_watch_group(contact=contact, group_name=watch_group_name):
        return
    if '[@ME]' in content:
        bot.SendTo(contact, "@"+member+" at命令未启用")
    dialog = dialog_clearify(content)
    dialog = is_cmd(dialog)
    if dialog is None:
        print("不是预定会议室")
    else:
        yuding_info = member.name
        book_ornot = find_yuding(dialog)  # True False
        riqi = find_riqi(dialog)  # 2018-03-13

        start_time, end_time = find_shijian(dialog)  # 8:00, 9:15
        if start_time is None or end_time is None:
            print("没找到时间")
        else:
            print("获取时间完成", start_time, end_time)
            fangjian = find_fangjian(dialog)  # 和昌12楼小会议室
            print("获取房间名完成")
            # 表格文件对象
            excel_file = get_excel_file(filename=excel_file_name)  # 检查是否被占用，，，
            print("获取文件", excel_file)
            excel_sheet = get_excel_sheet(riqi=riqi, file=excel_file)
            print("获取sheet", excel_sheet)
            # excel_column = get_excel_column(fangjian)  # todo 没有就新建 找的时候在第一行寻找 异常处理
            column0 = 2
            excel_column = fangjian + column0
            print("获取列", excel_column)
            excel_date_row = get_excel_row(sheet=excel_sheet, today=riqi)  # 下一行是8:00:00
            print("获取行", excel_date_row)
            delta_row_start, delta_row_end = get_dtime(start_time, end_time)
            print("获取行区段", delta_row_start, delta_row_end)
            deal_book(sheet=excel_sheet, start=excel_date_row + delta_row_start,
                      end=excel_date_row + delta_row_end, column=excel_column,
                      info=yuding_info, book=book_ornot, bot=bot, contact=contact)

            # 存储表格文件
            excel_file.save(filename=excel_file_name)
            print("\n")



