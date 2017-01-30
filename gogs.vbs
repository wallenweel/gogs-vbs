Dim currApp, currAppName

currAppName = Left(WScript.ScriptName, InStr(WScript.ScriptName, ".") - 1)
currApp = currAppName & ".exe"

' Entry of this App
Main(currApp)

Public Sub Main(ByVal sAppName)
    Dim fso, CurDir
    Set fso = CreateObject("Scripting.Filesystemobject")
    CurDir = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\"))

    ' Check app's process
    If RuningCount(sAppName, "") > 0 Then
        
        ' Check whether has log file
        If Not fso.FileExists(CurDir & "LOG.TXT") Then
        
            ' Kill app's process
            ' Call CloseProcess(sAppName, "")
            
            ' Notice message
            MsgBox currApp, 0, Time
        End If
    
    Else

        ' Notice message
        MsgBox currApp, 4, Time
        
    End If
    
End Sub

' Public Function FunctionName(ParameterList) As ReturnType
'     Try
        
'     Catch ex As Exception
'     End Try
'     Return ReturnValue
' End Function
' Count app process
' Eg: If RuningCount("cmd.exe", "") > 0
' Eg: If RuningCount("cmd.exe", "c:\0.bat") > 1
Public Function RuningCount(ByVal sAppName, ByVal sAppPath) 
    On Error Resume Next
    Dim objItem, i:    i = 0
    For Each objItem In GetObject("winmgmts:\\.\root\cimv2").instances_
        If LCase(objItem.Name) = LCase(sAppName) Then
            If sAppPath = "" Or InStr(1, objItem.CommandLine, sAppPath, vbTextCompare) Then i = i + 1
        End If
    Next
    RuningCount = i
End Function


' Close app process
' Eg: Call CloseProcess("cmd.exe", "")
' Eg: Call CloseProcess("cmd.exe", "c:\0.bat")
Sub CloseProcess(ByVal sAppName, ByVal sAppPath)
    On Error Resume Next
    Dim objItem
    For Each objItem In GetObject("winmgmts:\\.\root\cimv2").instances_
        If LCase(objItem.Name) = LCase(sAppName) Then
            If sAppPath = "" Or InStr(1, objItem.CommandLine, sAppPath, vbTextCompare) Then objItem.Terminate
        End If
    Next
End Sub