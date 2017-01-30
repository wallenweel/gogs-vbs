Set wim = GetObject("winmgmts:")
Set wso = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.filesystemobject")

Dim appName, postFix, argv, html

' Custom app Name
appName = ""

' Custom app postfix
postFix = ".exe"

argv = "web"

html = "index.html"

' Call main sub
Main

' Main
Sub Main()
    Dim xName, xPath, xCmd, xSql
    Dim currDir:currDir = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\"))

    xName = scriptName(appName, postFix)
    xPath = currDir & xName
    xSql = processSQL(xName, xPath)

    If Not fso.fileExists(xPath) Then
        MsgBox "No App: [" & xName & "], Will Exit!"
    Else
        Dim hasRan:hasRan = wim.execQuery(xSql).count
        ' hasRan = 0
        If hasRan = 1 Then
            Dim sts

            sts = MsgBox("[" & xName & "] is Running!", 2, "How do you want to do?")

            If sts = 3 Then Call terminateProcess(xSql, xName)
            If sts = 4 Then 
                Call terminateProcess(xSql, xName)
                wso.run xPath & " " & argv, 0
            End If
        ElseIf hasRan = 2 Then
            MsgBox ""
        Else
            wso.run xPath & " " & argv, 0
            ' LaunchGUI(currDir & html)
        End If
    End If
    
End Sub

Private Function terminateProcess(sql, name)
    For Each objItem In wim.execQuery(sql)
        If LCase(objItem.Name) = LCase(name) Then
            objItem.terminate
        End If
    Next
End Function

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

' Generate Process SQL
Private Function processSQL(name, path)
    Dim r

    r = "Select * From Win32_Process Where Name='{$1}'"
    ' r = "Select * From Win32_Process Where Name='{$1}' And CommandLine Like '%{$2}%'"
    ' path = Replace(path, "\", "\\")
    ' r = Replace(r, "{$2}", path)
    r = Replace(r, "{$1}", name)
    
    processSQL = r
End Function

Private Function LaunchGUI(path)
    Set IE = WScript.createObject("InternetExplorer.Application", "event_")
    
    IE.menubar = 0
    IE.addressbar = 0
    IE.toolbar = 0
    IE.statusbar = 0
    IE.width = 400
    IE.height = 400
    IE.resizable = 0
    ' IE.navigate "http://www.baidu.com"
    ' IE.navigate "file://" & path
    IE.navigate "about:blank"
    
    Do  
        WScript.Sleep 200  
    Loop Until IE.readyState = 4

    IE.left = Fix((IE.document.parentwindow.screen.availwidth - IE.width) / 2)
    IE.top = Fix((IE.document.parentwindow.screen.availheight - IE.height) / 2)
    IE.visible = 1

    ' Set http = CreateObject("Msxml2.ServerXMLHTTP")
    ' http.open "GET" path, False
    ' http.send
    ' sHtml = http.responseText

    IE.document.write sHtml
    
    MsgBox IE.document

    ' Dim stuff
    ' Set fso = CreateObject("Scripting.filesystemobject")
    ' Set file = fso.openTextFile(path)
    ' stuff = file.readAll
    ' IE.document.write stuff
    ' file.close
    

    LaunchGUI = IE
End Function