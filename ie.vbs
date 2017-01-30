set ie=wscript.createobject("internetexplorer.application","event_") '创建ie对象'
ie.menubar=0 '取消菜单栏'
ie.addressbar=0 '取消地址栏'
ie.toolbar=0 '取消工具栏'
ie.statusbar=0 '取消状态栏'
ie.width=400 '宽400'
ie.height=400 '高400'
ie.resizable=0 '不允许用户改变窗口大小'
ie.navigate "about:blank" '打开空白页面'
ie.left=fix((ie.document.parentwindow.screen.availwidth-ie.width)/2) '水平居中'
ie.top=fix((ie.document.parentwindow.screen.availheight-ie.height)/2) '垂直居中'
ie.visible=1 '窗口可见'

with ie.document '以下调用document.write方法，'
.write "<html><body bgcolor=#dddddd scroll=no>" '写一段html到ie窗口中。'
.write "<h2 align=center>远程清除系统日志</h2><br>"
.write "<p>目标IP<input id=ip type=text size=15>" '也可以用navigate方法直接打开一'
.write "<p>用户名：<input id=user type=text size=30>" '个html文件，效果是一样的。'
.write "<p>密码：　<input id=pass type=password size=30>"
.write "<p align=center>类型" '不仅是input对象，所有DHTML支持'
.write "<input id=app type=checkbox>应用程序 " '的对象及其属性、方法都可以使用。'
.write "<input id=sys type=checkbox>系统 "
.write "<input id=sec type=checkbox>安全" '访问这些对象的办法和网页中访问'
.write "<p align=center><br>" '框架内对象是类似的。'
.write "<input id=confirm type=button value=确定> "
.write "<input id=cancel type=button value=取消>"
.write "</body></html>"
end with

dim wmi '显式定义一个全局变量'
set wnd=ie.document.parentwindow '设置wnd为窗口对象'
set id=ie.document.all '设置id为document中全部对象的集合'
id.confirm.onclick=getref("confirm") '设置点击"确定"按钮时的处理函数'
id.cancel.onclick=getref("cancel") '设置点击"取消"按钮时的处理函数'

do while true '由于ie对象支持事件，所以相应的，'
wscript.sleep 200 '脚本以无限循环来等待各种事件。'
loop

sub event_onquit 'ie退出事件处理过程'
wscript.quit '当ie退出时，脚本也退出'
end sub

sub cancel '"取消"事件处理过程'
ie.quit '调用ie的quit方法，关闭IE窗口'
end sub '随后会触发event_onquit，于是脚本也退出了'

sub confirm '"确定"事件处理过程，这是关键'
with id
if .ip.value="" then .ip.value="." '空ip值则默认是对本地操作'
if not (.app.checked or .sys.checked or .sec.checked) then 'app等都是checkbox，通过检测其checked'
wnd.alert("至少选择一种日") '属性，来判断是否被选中。'
exit sub
end if
set lct=createobject("wbemscripting.swbemlocator") '创建服务器定位对象'
on error resume next '使脚本宿主忽略非致命错误'
set wmi=lct.connectserver(.ip.value,"root/cimv2",.user.value,.pass.value) '连接到root/cimv2名字空间'
if err.number then '自己捕捉错误并处理'
wnd.alert("连接WMI服务器失") '这里只是简单的显示“失败”'
err.clear
on error goto 0 '仍然让脚本宿主处理全部错误'
exit sub
end if
if .app.checked then clearlog "application" '清除每种选中的日志'
if .sys.checked then clearlog "system"
if .sec.checked then clearlog "security" '注意，在XP下有限制，不能清除安全日志'
wnd.alert("日志已清")
end with
end sub

sub clearlog(name)
wql="select * from Win32_NTEventLogFile where logfilename='"&name&"'"
set logs=wmi.execquery(wql) '注意，logs的成员不是每条日志，'
for each l in logs '而是指定日志的文件对象。'
if l.cleareventlog() then
wnd.alert("清除日志"&name&"时出错！")
ie.quit
wscript.quit
end if
next
end sub