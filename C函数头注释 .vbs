Set WshShell=WScript.CreateObject("WScript.Shell")

If  Msgbox("���л���Ӣ�����뷨��"& Chr(10) & "����[ȷ��]��,����궨λ������λ�ô���" & Chr(10) & "����[ȡ��]�˳�",65,"C����ͷע��") = 1 Then
    WScript.Sleep 3000
    WshShell.SendKeys "/*********************************************" & "{ENTER}"
    WshShell.SendKeys "* @brief  " & "{ENTER}"
    WshShell.SendKeys "* @param  " & "{ENTER}"
    WshShell.SendKeys "* @retval " & "{ENTER}"
    WshShell.SendKeys "**********************************************/" & "{ENTER}"
End If