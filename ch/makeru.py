

def bot_reply(bot, contact, str):
    _str = '机器人回复: '+str
    bot.SendTo(contact, _str)


watch_group_name = ['创客学院管理员社群',
                    'VR/AR/Unity3D/C#/创客',
                    # '创客学院VR/AR②群',
                    # 'Web前端/HTML/创客学院',
                    # '物联网_ARM_创客学院',
                    # '单片机/ARM/STM32/创客',
                    # '创客学院嵌入式ARM_STM32',
                    # 'Java/mysql/Oracle创客学院',
                    # '嵌入式/单片机/ARM/创客'
                    ]
ans_dic = {
    '创客学院官网': 'http://www.makeru.com.cn/ ',
    'VR教程': '先学c#和unity3d引擎 http://www.makeru.com.cn/roadmap/vr \n '
            '再学用unity3d的两个插件 \n'
            'https://assetstore.unity.com/packages/templates/systems/steamvr-plugin-32647 \n '
            'https://assetstore.unity.com/packages/tools/integration/vive-input-utility-64219 ',
    'AR教程': '高通vuforia https://developer.vuforia.com/downloads/sdk \n '
            '亮风台HiAR http://www.hiar.com.cn/doc-v1/main/home/ \n '
            'EasyAR https://www.easyar.cn/view/download.html \n '
            '苹果ARkit https://developer.apple.com/cn/arkit/ \n '
            '谷歌ARcore https://developer.apple.com/cn/arkit/ \n '
            '选一个吧',
    'web教程': 'http://www.makeru.com.cn/roadmap/web ',
    '嵌入式教程': 'http://www.makeru.com.cn/roadmap/emb ',
    'stm32教程': 'http://www.makeru.com.cn/search?q=stm32 ',
    '物联网教程': 'http://www.makeru.com.cn/roadmap/iot ',
    'arm教程': 'http://www.makeru.com.cn/search?q=arm ',
    'java教程': 'http://www.makeru.com.cn/roadmap/javaee ',
    '安卓教程': 'http://www.makeru.com.cn/roadmap/android ',
    '免费课程': 'http://www.makeru.com.cn/course/library?isPay=0 ',
    '直播课': 'http://www.makeru.com.cn/live/library ',
    'unity下载': 'https://unity3d.com/cn/get-unity/download/archive '
}


def onQQMessage(bot, contact, member, content):
    # 避免机器人自嗨 机器人发言请注意加上这个字符串
    if '机器人回复' in content:
        return
    # 监视制定的群
    if not (contact.nick in watch_group_name):
        return
    for k, v in ans_dic.items():
        if k in content:
            bot_reply(bot, contact, v)
    if '所有关键词' in content:
        s = ' '.join(ans_dic.keys())
        bot_reply(bot, contact, s)
        return
