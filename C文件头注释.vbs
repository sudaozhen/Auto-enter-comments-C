Set WshShell=WScript.CreateObject("WScript.Shell")

If  Msgbox("请切换成英文输入法！"& Chr(10) & "单击[确定]后,将光标定位到插入位置处。" & Chr(10) & "单击[取消]退出",65,"C文件头注释") = 1 Then

    Dim d
    d = Date()

    Dim author
    author = "Su DaoZhen"


    WScript.Sleep 3000
    WshShell.SendKeys "/**" & "{ENTER}"
    WshShell.SendKeys "  ******************************************************************************" & "{ENTER}"
    WshShell.SendKeys "  * @file  " & "{ENTER}"
    WshShell.SendKeys "  * @author  " & author & "{ENTER}"
    WshShell.SendKeys "  * @version  V1.0.0 " & "{ENTER}"
    WshShell.SendKeys "  * @date  " & date & "{ENTER}"
    WshShell.SendKeys "  * @brief  " & "{ENTER}"
    WshShell.SendKeys "  ******************************************************************************" & "{ENTER}"
    WshShell.SendKeys "  * @attention" & "{ENTER}" 
    WshShell.SendKeys "  * " & "{ENTER}"
    WshShell.SendKeys "  * " & "{(}" & "C" & "{)}" & "COPYRIGHT "  & author & "{ENTER}"
    WshShell.SendKeys "  ******************************************************************************" & "{ENTER}"
    WshShell.SendKeys "*/" & "{ENTER}"
    WshShell.SendKeys "{ENTER}" & "{ENTER}" & "{ENTER}" & "{ENTER}" & "{ENTER}"
    WshShell.SendKeys "/*******************" & "{(}" & "C" & "{)}" & " COPYRIGHT " & author & " *********END OF FILE*********/"
    WshShell.SendKeys "{ENTER}"
End If