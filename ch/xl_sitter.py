from meeting_tools import *


watch_group_name = ["0.0", '张志琳、刘文、李豪、', 'SIPPR 智能与信息12楼', '智能与信息工程中心']


def onQQMessage(bot, contact, member, content):
    # 避免机器人自嗨 机器人发言请注意加上这个字符串
    if '机器人回复' in content:
        return
    # 监视制定的群
    if not my_watch_group(contact=contact, group_name=watch_group_name):
        return

    if '[@ME]' in content:
        if 'start' in content:
            bot.Plug('xl')
          