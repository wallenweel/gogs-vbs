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
        hasRan = 0
        If hasRan = 1 Then
            Dim sts

            sts = MsgBox("How do you want to do?", 2, "[" & xName & "] is Running!")

            If sts = 3 Then Call terminateProcess(xSql, xName)
            If sts = 4 Then 
                Call terminateProcess(xSql, xName)
                wso.run xPath & " " & argv, 0
            End If
        ElseIf hasRan = 2 Then
            MsgBox ""
        Else
            ' wso.run xPath & " " & argv, 0
            LaunchGUI(currDir & html)
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
    r = "Select * From Win32_Process Where Name='{$1}'" ' And CommandLine Like '%{$2}%'"
    ' path = Replace(path, "\", "\\")
    ' r = Replace(r, "{$2}", path)
    r = Replace(r, "{$1}", name)
    
    processSQL = r
End Function

Private Function LaunchGUI(path)
    Set IE = WScript.createObject("InternetExplorer.Application", "event_")
    
    With IE
        .menubar = 0
        .addressbar = 0
        .toolbar = 0
        .statusbar = 0
        .width = 320
        .height = 560
        .resizable = 0
        .navigate "about:blank"
    
        Do
            WScript.Sleep 200
        Loop Until .readyState = 4
    
        .left = Fix((.document.parentwindow.screen.availwidth - .width) / 2)
        .top = Fix((.document.parentwindow.screen.availheight - .height) / 2)
        .document.write readFile(path, "utf-8")
        .visible = 1

    End With

    LaunchGUI = IE
End Function

Private Function getText(url)
    Set http = CreateObject("Msxml2.ServerXMLHTTP")

    http.open "GET", url, False
    http.send
    
    getText = http.responseText
End Function

Private Function readFile(path, charset)
    Dim Str
    Set Stuff = CreateObject("ADODB.Stream")

    With Stuff
        .type = 2
        .mode = 3
        .charset = charset
        .open
        .loadFromFile path

        Str = .readtext

        .close
    End With
    Set Stuff = Nothing

    ReadFile = Str
End Function

Private Function writeFile (content, file, charset)
    Set Stuff = CreateObject("ADODB.Stream")
    With Stuff
        .type = 2
        .mode = 3
        .charSet = charset
        .open
        .writeText content
        .saveToFile file, 2
        .flush
        .close
    End With
    Set Stuff = Nothing
End Function