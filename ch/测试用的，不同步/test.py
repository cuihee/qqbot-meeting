from ch.meeting_tools import *


dialog = '预约29号下午四点半和昌13楼会议室'
d = dialog_clearify(dialog)
print(d)
print(find_yuding(d))
print(find_riqi(d))
print(find_shijian(d))
print(find_fangjian(d))
print(get_meetingrooms_names()[find_fangjian(d)])
#？？？忽略了呀，怎么会上传

