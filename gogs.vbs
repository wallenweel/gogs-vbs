Set wim = GetObject("winmgmts:")
Set wso = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.filesystemobject")
Set IE  = WScript.createObject("InternetExplorer.Application", "event_")

Dim appName, postFix, autoStart, argv, html

' Custom app Name e.g. cmd
appName = ""

' Custom app postfix
postFix = ".exe"

' Specified auto start app, 0|1
autoStart = 0

' Specified argv
argv = "web"

' GUI page file, path is relative to the vbs script
html = "index.html"

Dim xDir, xName, xPath, xCmd, xSql, hasRan

' Call main sub
Main

' Main
Sub Main
    xDir  = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\"))
    xName = scriptName(appName, postFix)
    xPath = xDir & xName
    xCmd  = xName & " " & argv
    xSql  = processSQL(xName, xPath)

    If Not fso.fileExists(xPath) Then
        MsgBox "No App: [" & xName & "], Will Exit!"
    Else
        hasRan = wim.execQuery(xSql).count

        If html = "" Then
            If hasRan = 1 Then
                Dim sts
                sts = MsgBox("HOW YOU DO?", 2, "[" & xName & "] is Running!")

                If sts = 3 Then app_stop
                If sts = 4 Then app_restart
            ElseIf hasRan = 2 Then
                MsgBox ""
            Else
                app_start
                wso.popup "APP HAS LAUNCHED...", 5, "INFO", 0
            End If
        Else
            If autoStart = 1 Then app_start
            Call LaunchGUI(xDir & html)
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
    ' Set IE = WScript.createObject("InternetExplorer.Application", "event_")
    
    With IE
        .menubar = 0
        .addressbar = 0
        .toolbar = 0
        .statusbar = 0
        .width = 320
        .height = 568
        .resizable = 0
        .silent = 0
        .navigate "about:blank"
    
        Do
            WScript.sleep 200
        Loop Until .readyState = 4
    
        .left = Fix((.document.parentwindow.screen.availwidth - .width) / 2)
        .top = Fix((.document.parentwindow.screen.availheight - .height) / 2)
        .visible = 1

    End With

    With IE.document
        If Not fso.fileExists(path) Then
            .write "<html>"
            .write "<body style=text-align:center;padding-top:50%;>"
            .write "<button id=start>START</button>"
            .write "<button id=restart>RESTART</button>"
            .write "<button id=stop>STOP</button>"
            .write "</body>"
            .write "</html>"
        Else
            .write readFile(path, "utf-8")
        End If
        
        Set startBtn = .querySelector("button#start")
        Set restartBtn = .querySelector("button#restart")
        Set stopBtn = .querySelector("button#stop")
        Set goEntry = .querySelector("a#goEntry")
        Set goLog = .querySelector("a#goLog")
    End With

    startBtn.onclick = getRef("app_start")
    restartBtn.onclick = getRef("app_restart")
    stopBtn.onclick = getRef("app_stop")
    goEntry.onclick = getRef("app_go")
    goLog.onclick = getRef("app_go")

    If hasRan > 0 Then Call addSamp(xCmd)
        
    Do While true
        Call refreshStatus()
        WScript.sleep 800
    Loop

    LaunchGUI = IE
End Function

Private Function refreshStatus()
    Dim hasRan:hasRan = wim.execQuery(xSql).count
    on error resume next

    Set id = IE.document.all

    If hasRan = 0 Then
        id.start.disabled = null
        id.restart.disabled = true
        id.stop.disabled = true
    Else
        id.start.disabled = true
        id.restart.disabled = null
        id.stop.disabled = null
    End If
    
End Function

Public Sub event_onquit
    WScript.quit(0)
End Sub

Private Sub app_go(ev)
    With IE.document
        .body.className = ev.currentTarget.className
        If .body.className = "log" Then
            .querySelector("pre#log").innerHtml = readFile((xDir & "\readme_zh.md"), "utf-8")
        End If
    End With
End Sub

Private Sub app_start
    wso.run xCmd, 0
    Call addSamp(xCmd)
End Sub

Private Sub app_restart(ev)
    Call terminateProcess(xSql, xName)
    wso.run xCmd, 0
End Sub

Private Sub app_stop
    Call addSamp("")
    Call terminateProcess(xSql, xName)
End Sub

Private Function addSamp(str)
    Set obj = IE.document.querySelector("blockquote.comment > code")
    Dim old:old = obj.innerHtml
    obj.innerHtml = "<samp>" & trim(xDir, "\") & "&gt;&nbsp;" & str & "</samp>"
End Function

Private Function trim(str, char)
    trim = Left(str, InStrRev(str, char) - 1)
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