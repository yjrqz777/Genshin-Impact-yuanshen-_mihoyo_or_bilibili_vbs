a=msgbox("来我空间看看吧?",vbOKCancel)

if a=vbok then

Set Sh = WScript.CreateObject("WScript.Shell")

Sh.Run "http://t.csdnimg.cn/GU1z3", 3

end if