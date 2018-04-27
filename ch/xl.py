from meeting_tools import *


watch_group_name = ["0.0", 'SIPPR 智能与信息12楼', '智能与信息工程中心']
watch_buddy_name = ['崔鹤', '乔辉', '游冰']
excel_file_name = "testMeeting.xlsx"  # 新建的sheet第一天不太对，但是不影响使用
cmd_list = ['效率助手下载地址', '查询今天会议室预订情况', '查询明天会议室预订情况', 'help',  # 0 1 2 3
            '顺丰的联系方式', 'stop', 'watch_group_name']
# 关键词和触发的行为
cmd_dic = {
    '效率助手下载地址': '',
    '查询今天会议室预订情况': '',
    '查询明天会议室预订情况': '',
    'help': '',
    '顺丰的联系方式': '',
    'stop': '',
    'watch_group_name': ''
}


def onQQMessage(bot, contact, member, content):
    # 避免机器人自嗨 机器人发言请注意加上这个字符串
    if '机器人回复' in content:
        return
    # 监视制定的群
    if not my_watch_group(contact=contact, group_name=watch_group_name):
        return
    # todo 用dic替换list
    for k, v in cmd_dic.items():
        if k in content:
            if isinstance(cmd_dic[k], type(func())):
                # 是函数 就执行函数
                pass
            else:
                # 不是函数 直接输出
                pass
            # return
    if cmd_list[0] in content:
        bot.SendTo(contact, "机器人回复 eepm.sippr.cn ")
        return
    if cmd_list[1] in content:
        report_info = ask_info(excel_file_name, (datetime.date.today() + datetime.timedelta(days=0)).__str__())
        bot.SendTo(contact, '机器人回复 \n'+'\n'.join(report_info))
        return  # 避免对查询语句进行预定会议室
    if cmd_list[2] in content:
        report_info = ask_info(excel_file_name, (datetime.date.today()+datetime.timedelta(days=1)).__str__())
        bot.SendTo(contact, '机器人回复 \n'+'\n'.join(report_info))
        return  # 避免对查询语句进行预定会议室
    if '[@ME]' in content and cmd_list[3] in content:
        bot.SendTo(contact, '机器人回复 下列引号内的命令采用精确匹配。' + cmd_list.__str__())
        return
    if cmd_list[4] in content:
        bot.SendTo(contact, '机器人回复 顺丰 15936240735')
        return
    if '[@ME]' in content and cmd_list[5] in content:
        bot.SendTo(contact, '机器人回复 已停止')
        # bot.Stop()
        bot.Unplug('xl')
        return
    if '[@ME]' in content[:5]:
        bot.SendTo(contact, "机器人回复 只要你at我，我就回复这一句\n用来查看插件是否启用中")
        return
    if cmd_list[6] in content:
        bot.SendTo(contact, "机器人回复 "+watch_group_name.__str__())
        return

    dialog = dialog_clearify(content)
    dialog = is_cmd(dialog)
    if len(dialog) < 1:
        print("不是预定会议室")
        return
    else:
        yuding_info = member.name + ' 群"' + contact.nick + '" ' + datetime.datetime.today().__str__()[:-7] \
                      + '            '
        book_ornot = find_yuding(dialog)  # True False
        riqi = find_riqi(dialog)  # 2018-03-13
        # print("获取日期:", riqi)
        start_time, end_time = find_shijian(dialog)  # 8:00, 9:15
        # print("获取时间:", start_time, end_time)
        fangjian = find_fangjian(dialog)  # 和昌12楼小会议室
        # print("获取房间名:", get_meetingrooms_names()[fangjian])

        yuding_info = yuding_info + riqi + ' '
        yuding_info = yuding_info + get_meetingrooms_names()[fangjian] + ' '
        yuding_info = yuding_info + start_time + '-' + end_time + ' '
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

        # ==============================================注意=======================================================
        # 不管从哪个群里获取的预订信息都反馈到0.0群里面！！！！！！！！！！！！！！！！
        bl = bot.List('group', '0.0')
        b = None
        if bl:
            b = bl[0]
            # bot.SendTo(b, 'test')

        deal_book(sheet=excel_sheet, start=excel_date_row + delta_row_start,
                  end=excel_date_row + delta_row_end, column=excel_column,
                  info=yuding_info, book=book_ornot, bot=bot, contact=b,  # 注意这里的contact!!!!!!!!!!!!!!!!!!!!!
                  member=member)

        # 存储表格文件
        excel_file.save(filename=excel_file_name)
        print("\n")


def onPlug(bot):
    s = bot.Plugins()
    if 'xl_sitter' in s:
        pass
    else:
        bot.Plug('xl_sitter')


def func():
    pass

