Set WshShell=WScript.CreateObject("WScript.Shell")

If  Msgbox("���л���Ӣ�����뷨��"& Chr(10) & "����[ȷ��]��,����궨λ������λ�ô���" & Chr(10) & "����[ȡ��]�˳�",65,"C�ļ�ͷע��") = 1 Then

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