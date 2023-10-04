
# 电脑原神B服官服切换脚本（部分重写，此后不必每个版本更新，有问题请反馈）
# Genshin-Impact-yuanshen-_mihoyo_or_bilibili_vbs
## **电脑原神（此版本无需跟随更新）**
![在这里插入图片描述](https://img-blog.csdnimg.cn/427b5338b04445fa9abd0a562e46fed2.png)
==此版本无需跟随更新==
==此版本无需跟随更新==
==最终版，有问题请反馈==

## 最新脚本源码移步github

[https://github.com/yjrqz777/Genshin-Impact-yuanshen-_mihoyo_or_bilibili_vbs](https://github.com/yjrqz777/Genshin-Impact-yuanshen-_mihoyo_or_bilibili_vbs)

## 私服脚本移步

[私服点我](https://blog.csdn.net/weixin_51681760/article/details/124843316)



=本人原创！==

**有问题留言吧，QQ加的人太多了，问题太杂**
简单问题大家一起解决
链接：[点我跳转](https://yjrqz.lanzoui.com/b01oi27oh)
==**https://yjrqz.lanzoui.com/b01oi27oh
密码:8za8**==
==**特供版请不要主动使用、特供版请不要主动使用、特供版请不要主动使用**==
先放原图：
![在这里插入图片描述](https://img-blog.csdnimg.cn/278c35f7f8ad464db397293d25c42cbf.png?x-oss-process=image/watermark,type_d3F5LXplbmhlaQ,shadow_50,text_Q1NETiBA5byC5aKD5YWl5L616ICF,size_20,color_FFFFFF,t_70,g_se,x_16#pic_center)



好了；不再赘述了，简单说一下原理：
  第一：配置文件，我的脚本是通过修改配置文件，来切换服务器启动的
  第二：我用的是微软()自带脚本工具后缀为==.vbs==
  第三：简单告知一下配置文件是后缀为==.ini==的文件
  ![ ](https://img-blog.csdnimg.cn/a13eb8b78d4145578f0ead87cc03d0a0.png#pic_center)
源码如下：

```c
'time：2021.8.6-18.54--2021.8.6-19.34
'author:YJRQZ777

dim fz
set fz = CreateObject("Scripting.FileSystemObject")
call fz.CopyFile("config.ini", "config.ini7") '两个参数的文件名部分可以不同
set fz = nothing


dim fso, f
set fso = CreateObject("Scripting.FileSystemObject")
set f = fso.OpenTextFile("config.ini", 2, false) '第二个参数 2 表示重写，如果是 8 表示追加
'f.Write("写入内容")
'f.WriteLine("写入内容并换行")
'f.WriteBlankLines(3) '写入三个空白行（相当于在文本编辑器中按三次回车）


msgbox "作者通行证：503250004"
msgbox "B服选择:1!"&chr(10)&"官服选择:2!"
dim a
a=0
do while a=0
a = inputbox ("b服选择:1!"&chr(10)&"官服选择:2!"&chr(10)&"请不要输入其他东西！"&chr(10)&"有问题联系qq:3210551161"&chr(10)&"后缀为.ini7的是备份文件."&chr(10)&"出现其他问题可重命名（删除后缀上的“7”）文件恢复备份","电脑原神服务器修改1.0")
Loop
IF a=1 Then 
f.WriteLine("[General]")
f.WriteLine("channel=14")
f.WriteLine("cps=bilibili")
f.WriteLine("sub_channel=0")
f.WriteLine("game_version=2.0.0")
f.WriteLine("sdk_version=")
msgbox "成功修改为B服"
ElseIf a=2 Then 
f.WriteLine("[General]")
f.WriteLine("channel=1")
f.WriteLine("cps=mihoyo")
f.WriteLine("sub_channel=1")
f.WriteLine("game_version=2.0.0")
f.WriteLine("sdk_version=")
msgbox "成功修改为官服"
else msgbox "我不理解！"
End If
msgbox "感谢使用！"
f.Close()
set f = nothing
set fso = nothing
```
