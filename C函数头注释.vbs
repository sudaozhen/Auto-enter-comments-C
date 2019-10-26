Set WshShell=WScript.CreateObject("WScript.Shell")

If  Msgbox("请切换成英文输入法！"& Chr(10) & "单击[确定]后,将光标定位到插入位置处。" & Chr(10) & "单击[取消]退出",65,"C函数头注释") = 1 Then
    WScript.Sleep 3000
    WshShell.SendKeys "/*********************************************" & "{ENTER}"
    WshShell.SendKeys "* @brief  " & "{ENTER}"
    WshShell.SendKeys "* @param  " & "{ENTER}"
    WshShell.SendKeys "* @retval " & "{ENTER}"
    WshShell.SendKeys "**********************************************/" & "{ENTER}"
End If
