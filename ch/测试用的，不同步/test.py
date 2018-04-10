from ch.meeting_tools import *

yuding_info = 'renmingzi' + ' 群"' + 'qunmingzi' + '" ' + datetime.datetime.today().__str__()[:-7] + ' 预定的'
print(yuding_info)
riqi = '2018-04-10'
yuding_info = yuding_info + riqi + ' '
start_time = '18:00'
end_time = '19:15'
yuding_info = yuding_info + start_time + '-' + end_time + ' '
fangjian = 0
yuding_info = yuding_info + fangjian.__str__() + ' ' + get_meetingrooms_names()[fangjian] + ' '
print(yuding_info)
print(yuding_info[-32:])