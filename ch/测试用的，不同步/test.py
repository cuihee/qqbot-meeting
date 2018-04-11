from ch.meeting_tools import *


dialog = '预定会议室'
if dialog.find("会议室") > -1:
    if dialog.find("预") > -1 or dialog.find("订") > -1 or dialog.find("定") > -1:
        print(dialog)
if "会议室" in dialog:
    if "预" in dialog or "订" in dialog or "定" in dialog > -1:
        print(dialog)
