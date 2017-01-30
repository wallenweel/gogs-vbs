Dim appName, postFix

' Custom app Name
appName = ""

' Custom app postfix
postFix = ".exe"

' Call main sub
Main

' Main
Sub Main()
    Dim wim, wso, fso
    Dim xName, xPath, xCmd, xSql
    Dim currDir:currDir = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\"))

    xName = scriptName(appName, postFix)
    xPath = currDir & xName
    xSql = processSQL(xName, xPath)
    
    Set wim = GetObject("winmgmts:")
    Set wso = CreateObject("WScript.Shell")
    Set fso = CreateObject("Scripting.filesystemobject")

    If Not fso.fileExists(xPath) Then
        MsgBox "No App: [" & xName & "], Will Exit!"
    Else
        MsgBox wim.execQuery(xSql).count
    End If
    
End Sub

' Generate Script Name
' @param {String} name  App name e.g. "cmd"
' @param {String} post  App postfix with dot e.g. ".exe"
' @return {String} App full name e.g. "cmd.exe"
Private Function scriptName(name, post)
    Dim r

    r = Left(WScript.scriptName, InStr(WScript.scriptName, ".") - 1)

    If name = "" Then 
        r = r & post
    Else 
        r = name & post
    End If

    scriptName = r
End Function

Public Function processSQL(name, path)
    Dim r

    r = "Select * From Win32_Process Where Name='{$1}' And CommandLine Like '%{$2}%'"
    path = Replace(path, "\", "\\")
    r = Replace(r, "{$1}", name)
    r = Replace(r, "{$2}", path)
    
    processSQL = r
End Function