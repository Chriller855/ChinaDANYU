<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!-- #include file="inc_config.asp" -->
<%
'Ȩ�޼��

'�ӳ��ű���ʱʱ��
server.ScriptTimeout =240

'��ȡ����
offtime=Request("offtime")
if (not isdate(offtime)) then Response.end() 'Response.Redirect "help.asp?error=����ȷ��дҪ�������ݵĽ�ֹ���ڡ�"

'�������ݶ���
Set rs = Server.CreateObject("ADODB.Recordset")

set bakconn=server.createobject("adodb.connection")
bakDBPath = Server.MapPath(bakconnpath)
bakconn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & bakDBPath
Set bakrs = Server.CreateObject("ADODB.Recordset")

'====================
'    ������·��Ϣ
'====================

'����������ȡ��·��Ϣ
sql="select count(id) as tcid,vcome from view where vtime<datevalue('" & offtime & "') and bakpage=0 group by vcome"
rs.Open sql,conn,1,1

do while not rs.EOF 
	'����ǰ��׷�ӵ��󱸿�
	sql="select * from stats where vtype='come' and vtitle='" & rs("vcome") & "'"
	bakrs.Open sql,bakconn,3,2
	if bakrs.EOF then	'����󱸿���û����һ��
		on error resume next
		bakrs.AddNew 
		bakrs("vtype")="come"
		bakrs("vtitle")=rs("vcome")
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
'  ���ݱ���ҳ����Ϣ
'====================

'����������ȡ����ҳ����Ϣ
sql="select count(id) as tcid,vpage from view where vtime<datevalue('" & offtime & "') and bakpage=0 group by vpage"
rs.Open sql,conn,1,1

do while not rs.EOF 
	'����ǰ��׷�ӵ��󱸿�
	sql="select * from stats where vtype='page' and vtitle='" & rs("vpage") & "'"
	bakrs.Open sql,bakconn,3,2
	if bakrs.EOF then	'����󱸿���û����һ��
		bakrs.AddNew 
		bakrs("vtype")="page"
		bakrs("vtitle")=rs("vpage")
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
conn.execute("update view set bakpage=1 where vtime<datevalue('" & offtime & "')")

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
<title><%=countname%>�����ݱ��ݣ����Ĳ�������ҳ����Ϣ</title>
<link rel="stylesheet" type="text/css" href="../../index.css">
</head>
<body topmargin=5 rightmargin=0 leftmargin=0 vlink=#000000 style="BACKGROUND-IMAGE: none" class=backq>
<br>
<p class="p1">ҳ����Ϣ������ɡ�
<p class="p1" align="right"><a href='bak_step5.asp?offtime=<%=offtime%>' target="_parent">��һ�� ���</a>&nbsp;
</body>
</html>
