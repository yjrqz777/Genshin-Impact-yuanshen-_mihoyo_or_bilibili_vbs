'time：2023.10.04
'author：YJRQZ777
'QQ群：155374244


dim PCGameSDK_Path

PCGameSDK_Path="./YuanShen_Data/Plugins/PCGameSDK.dll"

Function IsExitAFile(filespec)
        Dim fso
        Set fso=CreateObject("Scripting.FileSystemObject")        
        If fso.fileExists(filespec) Then         
        IsExitAFile=True        
        Else IsExitAFile=False        
        End If
End Function 

IF IsExitAFile(PCGameSDK_Path)=False Then 

y_n=msgbox("PCGameSDK.dll文件缺失"&chr(10)&"请前往这里下载所需文件"&chr(10)&"https://yjrqz.lanzoui.com/b01oi27oh 密码:8za8"&chr(10)&"移动到原神此路径/YuanShen_Data/Plugins/",vbCritical + vbYesNo,"错误")

' "PCGameSDK.dll文件缺失"&chr(10)&"请前往这里下载所需文件"&chr(10)&"https://yjrqz.lanzoui.com/b01oi27oh 密码:8za8"&chr(10)&"移动到原神此路径/YuanShen_Data/Plugins/"
if y_n then

Set Sh = WScript.CreateObject("WScript.Shell")
Sh.Run "http://t.csdnimg.cn/GU1z3", 3

end if


ElseIf IsExitAFile(PCGameSDK_Path)=True Then

Set fso=CreateObject("Scripting.Filesystemobject")
Set dic=CreateObject("Scripting.Dictionary")

infile=".\config.ini"
outfile=".\config.ini7"

dim input_num
input_num=0
do while input_num=0
input_num = inputbox ("b服选择:1!"&chr(10)&"官服选择:2!"&chr(10)&"请不要输入其他东西！"&chr(10)&"有问题qq群:155374244"&chr(10)&"后缀为.ini7的是备份文件."&chr(10)&"出现其他问题可重命名（删除后缀上的“7”）文件恢复备份","电脑原神服务器修改2.0")
Loop
IF input_num=1 Then 
dic.Add "2","channel=14"
dic.Add "3","cps=bilibili"
dic.Add "4","sub_channel=0"
msgbox "成功修改为B服"
ElseIf input_num=2 Then 
dic.Add "2","channel=1"
dic.Add "3","cps=mihoyo"
dic.Add "4","sub_channel=1"
msgbox "成功修改为官服"
else msgbox "我不理解！"
End If





Set f1=fso.OpenTextFile(infile,1)
Set f2=fso.CreateTextFile(outfile,2)
n=0
Do While f1.AtEndOfStream<>true
    n=n+1
    line=f1.ReadLine
    If dic.Exists(CStr(n)) Then
        f2.WriteLine dic.Item(CStr(n))
    Else
        f2.WriteLine line
    End If
Loop
f1.Close
f2.Close

dim fz
set fz = CreateObject("Scripting.FileSystemObject")
call fz.CopyFile(outfile, infile) '两个参数的文件名部分可以不同
set fz = nothing
msgbox "感谢使用！"&chr(10)&"此为升级版，不必每个版本更新，有问题请反馈"&chr(10)&"QQ群：155374244"

'msgbox info

End If
