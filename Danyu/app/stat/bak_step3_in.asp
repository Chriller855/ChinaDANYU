<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!-- #include file="inc_config.asp" -->
<%
'Ȩ�޼��

'�ӳ��ű���ʱʱ��
server.ScriptTimeout =240

'��ȡ����
offtime=Request("offtime")
if (not isdate(offtime)) then Response.Redirect "help.asp?error=����ȷ��дҪ�������ݵĽ�ֹ���ڡ�"

'�������ݶ���
Set rs = Server.CreateObject("ADODB.Recordset")

set bakconn=server.createobject("adodb.connection")
bakDBPath = Server.MapPath(bakconnpath)
bakconn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & bakDBPath
Set bakrs = Server.CreateObject("ADODB.Recordset")

'====================
'   �����������Ϣ
'====================

'����������ȡ�������Ϣ
sql="select count(id) as tcid,vsoft from view where vtime<datevalue('" & offtime & "') and bakstats=0 group by vsoft"
rs.Open sql,conn,1,1

do while not rs.EOF 
	'����ǰ��׷�ӵ��󱸿�
	sql="select * from stats where vtype='soft' and vtitle='" & rs("vsoft") & "'"
	bakrs.Open sql,bakconn,3,2
	if bakrs.EOF then	'����󱸿���û����һ��
		bakrs.AddNew 
		bakrs("vtype")="soft"
		bakrs("vtitle")=rs("vsoft")
		bakrs("vdata")=rs("tcid")
		bakrs.Update 
	else		'����Ѿ�����һ���ˣ���׷������
		bakrs("vdata")=bakrs("vdata")+rs("tcid")
		bakrs.Update 
	end if
	bakrs.Close
	rs.MoveNext
loop
rs.Close

'====================
'  ���ݲ���ϵͳ��Ϣ
'====================

'����������ȡ����ϵͳ��Ϣ
sql="select count(id) as tcid,vOS from view where vtime<datevalue('" & offtime & "') and bakstats=0 group by vOS"
rs.Open sql,conn,1,1

do while not rs.EOF 
	'����ǰ��׷�ӵ��󱸿�
	sql="select * from stats where vtype='OS' and vtitle='" & rs("vOS") & "'"
	bakrs.Open sql,bakconn,3,2
	if bakrs.EOF then	'����󱸿���û����һ��
		bakrs.AddNew 
		bakrs("vtype")="OS"
		bakrs("vtitle")=rs("vOS")
		bakrs("vdata")=rs("tcid")
		bakrs.Update 
	else		'����Ѿ�����һ���ˣ���׷������
		bakrs("vdata")=bakrs("vdata")+rs("tcid")
		bakrs.Update
	end if
	bakrs.Close
	rs.MoveNext
loop
rs.Close


'====================
'  ������Ļ�����Ϣ
'====================

'����������ȡ��Ļ�����Ϣ
sql="select count(id) as tcid,vwidth from view where vtime<datevalue('" & offtime & "') and bakstats=0 group by vwidth"
rs.Open sql,conn,1,1

do while not rs.EOF 
	'����ǰ��׷�ӵ��󱸿�
	sql="select * from stats where vtype='width' and vtitle='" & rs("vwidth") & "'"
	bakrs.Open sql,bakconn,3,2
	if bakrs.EOF then	'����󱸿���û����һ��
		bakrs.AddNew 
		bakrs("vtype")="width"
		bakrs("vtitle")=rs("vwidth")
		bakrs("vdata")=rs("tcid")
		bakrs.Update 
	else		'����Ѿ�����һ���ˣ���׷������
		bakrs("vdata")=bakrs("vdata")+rs("tcid")
		bakrs.Update
	end if
	bakrs.Close
	rs.MoveNext
loop
rs.Close


'====================
'  ���ݷÿ͵�����Ϣ
'====================

'����������ȡ�����ߵ�����Ϣ
sql="select count(id) as tcid,vwhere from view where vtime<datevalue('" & offtime & "') and bakstats=0 group by vwhere"
rs.Open sql,conn,1,1

do while not rs.EOF 
	'����ǰ��׷�ӵ��󱸿�
	sql="select * from stats where vtype='where' and vtitle='" & rs("vwhere") & "'"
	bakrs.Open sql,bakconn,3,2
	if bakrs.EOF then	'����󱸿���û����һ��
		bakrs.AddNew 
		bakrs("vtype")="where"
		bakrs("vtitle")=rs("vwhere")
		bakrs("vdata")=rs("tcid")
		bakrs.Update 
	else		'����Ѿ�����һ���ˣ���׷������
		bakrs("vdata")=bakrs("vdata")+rs("tcid")
		bakrs.Update
	end if
	bakrs.Close
	rs.MoveNext
loop
rs.Close


'�����������ϱ�ǣ�ָ����Щ���ݵġ��ͻ�����Ϣ���Ѿ�����
conn.execute("update view set bakstats=1 where vtime<datevalue('" & offtime & "')")

set rs=nothing
conn.Close
set conn=nothing

set bakrs=nothing
bakconn.Close
set bakconn=nothing




%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<title><%=countname%>�����ݱ��ݣ������������ݿͻ�����Ϣ</title>
<link rel="stylesheet" type="text/css" href="../../index.css">
</head>
<body topmargin=5 rightmargin=0 leftmargin=0 vlink=#000000 style="BACKGROUND-IMAGE: none" class=backq>
<br>
<p class="p1">�ͻ�����Ϣ������ɡ�
<p class="p1" align="right"><a href='bak_step4.asp?offtime=<%=offtime%>' target="_parent">��һ�� ����ҳ����Ϣ ��ʼ</a>&nbsp;
</body>
</html>
