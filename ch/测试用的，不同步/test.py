from ch.meeting_tools import *

#excel_file_name = "testMeeting.xlsx"
#file = load_workbook(excel_file_name)
#sheet = get_excel_sheet('2018-04-02', file)
#print(sheet.cell(row=1, column=1).value)
#print(sheet.cell(row=1, column=2).value)
#print(sheet.cell(row=1, column=3).value)
#print(sheet.cell(row=1, column=4).value)
#print(sheet.cell(row=1, column=5).value)
#sheet.cell(row=1, column=1).value = None
#print(sheet.cell(row=1, column=1).value)
#
#file.save(excel_file_name)

d = '预订12楼大会议室，12点到14点'
d = dialog_clearify(d)
d = is_cmd(d)
print(find_riqi(d))
print(find_shijian(d)[0])
print(find_shijian(d)[1])
print(get_meetingrooms_names()[find_fangjian(d)])
