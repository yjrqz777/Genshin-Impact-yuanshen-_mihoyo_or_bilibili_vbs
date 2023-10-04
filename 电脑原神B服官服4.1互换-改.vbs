'time：2022.6.7


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


msgbox "qq群:155374244"&chr(10)&"B服选择:1!"&chr(10)&"官服选择:2!"
dim a
a=0
do while a=0
a = inputbox ("b服选择:1!"&chr(10)&"官服选择:2!"&chr(10)&"请不要输入其他东西！"&chr(10)&"有问题qq群:155374244"&chr(10)&"后缀为.ini7的是备份文件."&chr(10)&"出现其他问题可重命名（删除后缀上的“7”）文件恢复备份","电脑原神服务器修改1.0")
Loop
IF a=1 Then 
f.WriteLine("[General]")
f.WriteLine("channel=14")
f.WriteLine("cps=bilibili")
f.WriteLine("sub_channel=0")
f.WriteLine("game_version=4.1.0")
f.WriteLine("plugin_5_version=2.6.0")
f.WriteLine("plugin_sdk_version=3.5.0")
msgbox "成功修改为B服"
ElseIf a=2 Then 
f.WriteLine("[General]")
f.WriteLine("channel=1")
f.WriteLine("cps=mihoyo")
f.WriteLine("sub_channel=1")
f.WriteLine("game_version=4.1.0")
f.WriteLine("plugin_5_version=2.6.0")
f.WriteLine("plugin_sdk_version=3.5.0")
msgbox "成功修改为官服"
else msgbox "我不理解！"
End If
msgbox "感谢使用！"&chr(10)&"请注意每个版本必更新"&chr(10)&"https://yjrqz.lanzoui.com/b01oi27oh 密码:8za8"
f.Close()
set f = nothing
set fso = nothing