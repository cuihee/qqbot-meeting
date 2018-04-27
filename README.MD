本项目是基于qqbot[https://github.com/pandolia/qqbot]写的一个插件
（我写的脚本都在ch文件夹里面，ch文件夹都是我写的，其余文件都是qqbot项目的）

本项目和qqbot都是使用python实现（我用的是anaconda版 python3.6.3）

本项目的github地址：https://github.com/cuihee/qqbot-meeting 

本项目实现的功能：
在qqbot基础上，
维护一个excel表格，使之可以实现
·预定会议室，
·取消预定会议室，
·查看预订情况。

#安装步骤与使用方法
1 安装anaconda（python3.6）版本
（这个版本好像直接配置好了系统变量）

2 按照qqbot中的描述安装qqbot https://github.com/pandolia/qqbot

3 下载本项目中的ch文件夹

4 根据需要服务的qq群名称修改ch\xl.py中的代码

5 待写

6 将ch\xl.py放到qqbot的插件目录 我的是C:\Users\c2534\.qqbot-tmp\plugins

7 将ch\meeting_tools.py放到系统变量区域内 比如说D:\anaconda\Lib

8 启动一个qqbot，方法在qqbot项目里面有

9 成功启动后，qqbot默认加载了显示聊天信息的插件，所以在控制台可以看见登录qq的所有群的聊天信息

10 启用插件xl，方法也在qqbot项目里面

11 在服务的群中使用@命令或者预定会议室的命令触发机器人

#感谢qqbot及其所依赖项目
#反馈
在github中提issue或者发邮件给我c253406143@163.com （我不一定会看到）