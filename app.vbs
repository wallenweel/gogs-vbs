Main("gogs.exe")
Sub Main(ByVal sAppName)
    Dim fso, CurDir
    Set fso = CreateObject("Scripting.Filesystemobject")
    CurDir = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName,"\"))

    '检测CMD.exe进程是否存在
    If RuningCount(sAppName, "") > 0 Then
        
        ' 检测同目录下LOG.TXT有无生成
        If Not fso.FileExists(CurDir & "LOG.TXT") Then
        
            ' 结束 cmd.exe 进程
            Call CloseProcess(sAppName, "")
            
            ' 弹出提示
            ' MsgBox "程序遇到未知问题即将关闭，请重新运行本程序", 48, "友情提示"
            
        End If
    
    Else

        ' MsgBox "程序问题", 48, "友情提示"
        msgbox WScript
        
    End If
    
    ' 删除自己
    ' fso.DeleteFile WScript.ScriptFullName, True

End Sub


' 统计进程数
' Eg: If RuningCount("cmd.exe", "") > 0
' Eg: If RuningCount("cmd.exe", "c:\0.bat") > 1
Function RuningCount(ByVal sAppName, ByVal sAppPath)
    On Error Resume Next
    Dim objItem, i:    i = 0
    For Each objItem In GetObject("winmgmts:\\.\root\cimv2:win32_process").instances_
        If LCase(objItem.Name) = LCase(sAppName) Then
            If sAppPath = "" Or InStr(1,objItem.CommandLine,sAppPath,vbTextCompare) Then i = i + 1
        End If
    Next
    RuningCount = i
End Function


' ----------------------------------------------------------------------------------------------------
' 结束进程，指定程序、路径
' Eg: Call CloseProcess("cmd.exe", "")
' Eg: Call CloseProcess("cmd.exe", "c:\0.bat")
Sub CloseProcess(ByVal sAppName, ByVal sAppPath)
    On Error Resume Next
    Dim objItem
    For Each objItem In GetObject("winmgmts:\\.\root\cimv2:win32_process").instances_
        If LCase(objItem.Name) = LCase(sAppName) Then
            If sAppPath = "" Or InStr(1, objItem.CommandLine, sAppPath, vbTextCompare) Then objItem.Terminate
        End If
    Next
End Sub