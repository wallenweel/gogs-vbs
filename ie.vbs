set ie=wscript.createobject("internetexplorer.application","event_") '����ie����'
ie.menubar=0 'ȡ���˵���'
ie.addressbar=0 'ȡ����ַ��'
ie.toolbar=0 'ȡ��������'
ie.statusbar=0 'ȡ��״̬��'
ie.width=400 '��400'
ie.height=400 '��400'
ie.resizable=0 '�������û��ı䴰�ڴ�С'
ie.navigate "about:blank" '�򿪿հ�ҳ��'
ie.left=fix((ie.document.parentwindow.screen.availwidth-ie.width)/2) 'ˮƽ����'
ie.top=fix((ie.document.parentwindow.screen.availheight-ie.height)/2) '��ֱ����'
ie.visible=1 '���ڿɼ�'

with ie.document '���µ���document.write������'
.write "<html><body bgcolor=#dddddd scroll=no>" 'дһ��html��ie�����С�'
.write "<h2 align=center>Զ�����ϵͳ��־</h2><br>"
.write "<p>Ŀ��IP<input id=ip type=text size=15>" 'Ҳ������navigate����ֱ�Ӵ�һ'
.write "<p>�û�����<input id=user type=text size=30>" '��html�ļ���Ч����һ���ġ�'
.write "<p>���룺��<input id=pass type=password size=30>"
.write "<p align=center>����" '������input��������DHTML֧��'
.write "<input id=app type=checkbox>Ӧ�ó��� " '�Ķ��������ԡ�����������ʹ�á�'
.write "<input id=sys type=checkbox>ϵͳ "
.write "<input id=sec type=checkbox>��ȫ" '������Щ����İ취����ҳ�з���'
.write "<p align=center><br>" '����ڶ��������Ƶġ�'
.write "<input id=confirm type=button value=ȷ��> "
.write "<input id=cancel type=button value=ȡ��>"
.write "</body></html>"
end with

dim wmi '��ʽ����һ��ȫ�ֱ���'
set wnd=ie.document.parentwindow '����wndΪ���ڶ���'
set id=ie.document.all '����idΪdocument��ȫ������ļ���'
id.confirm.onclick=getref("confirm") '���õ��"ȷ��"��ťʱ�Ĵ�����'
id.cancel.onclick=getref("cancel") '���õ��"ȡ��"��ťʱ�Ĵ�����'

do while true '����ie����֧���¼���������Ӧ�ģ�'
wscript.sleep 200 '�ű�������ѭ�����ȴ������¼���'
loop

sub event_onquit 'ie�˳��¼��������'
wscript.quit '��ie�˳�ʱ���ű�Ҳ�˳�'
end sub

sub cancel '"ȡ��"�¼��������'
ie.quit '����ie��quit�������ر�IE����'
end sub '���ᴥ��event_onquit�����ǽű�Ҳ�˳���'

sub confirm '"ȷ��"�¼�������̣����ǹؼ�'
with id
if .ip.value="" then .ip.value="." '��ipֵ��Ĭ���ǶԱ��ز���'
if not (.app.checked or .sys.checked or .sec.checked) then 'app�ȶ���checkbox��ͨ�������checked'
wnd.alert("����ѡ��һ����") '���ԣ����ж��Ƿ�ѡ�С�'
exit sub
end if
set lct=createobject("wbemscripting.swbemlocator") '������������λ����'
on error resume next 'ʹ�ű��������Է���������'
set wmi=lct.connectserver(.ip.value,"root/cimv2",.user.value,.pass.value) '���ӵ�root/cimv2���ֿռ�'
if err.number then '�Լ���׽���󲢴���'
wnd.alert("����WMI������ʧ") '����ֻ�Ǽ򵥵���ʾ��ʧ�ܡ�'
err.clear
on error goto 0 '��Ȼ�ýű���������ȫ������'
exit sub
end if
if .app.checked then clearlog "application" '���ÿ��ѡ�е���־'
if .sys.checked then clearlog "system"
if .sec.checked then clearlog "security" 'ע�⣬��XP�������ƣ����������ȫ��־'
wnd.alert("��־����")
end with
end sub

sub clearlog(name)
wql="select * from Win32_NTEventLogFile where logfilename='"&name&"'"
set logs=wmi.execquery(wql) 'ע�⣬logs�ĳ�Ա����ÿ����־��'
for each l in logs '����ָ����־���ļ�����'
if l.cleareventlog() then
wnd.alert("�����־"&name&"ʱ����")
ie.quit
wscript.quit
end if
next
end sub