<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!-- #include file="inc_config.asp" -->
<%
'��ȡ����
offtime=Request("offtime")
if (not isdate(offtime)) then response.end() 'Response.Redirect "help.asp?error=����ȷ��дҪ�������ݵĽ�ֹ���ڡ�"

'�������ݶ���
'conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBPath

set bakconn=server.createobject("adodb.connection")
bakDBPath = Server.MapPath(bakconnpath)
bakconn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & bakDBPath
Set bakrs = Server.CreateObject("ADODB.Recordset")

'����������ȡÿ�շ�����
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select count(id) as tcid,datevalue(vtime) as tdate from view where vtime<datevalue('" & offtime & "') and bakdays=0 group by datevalue(vtime)"
rs.Open sql,conn,1,1

Set tmprs = Server.CreateObject("ADODB.Recordset")

bakok=false		'��ʼ������ɱ��Ϊδ���
for i=1 to 2
	if rs.EOF then
		bakok=true
		exit for
	end if
	'���㵱�յ�IP����
	sql="Select count(id) as abc from view where datevalue(vtime)=datevalue('" & rs("tdate") & "') group by vip"
	tmprs.Open sql,conn,1,1
	tcip	=tmprs.RecordCount
	tcid	=rs("tcid")
	tdate	=rs("tdate")
	tmprs.Close
	'Response.Write rs("tcid") & "," & tcip & "," & rs("tdate") & "<br>"
	'����ǰ��׷�ӵ��󱸿�
	sql="select * from days where datevalue(tdate)=datevalue('" & tdate & "')"
	bakrs.Open sql,bakconn,3,2
	if bakrs.EOF then	'����󱸿���û����һ��
	bakrs.AddNew 
	bakrs("tdate")=tdate
	bakrs("tview")=tcid
	bakrs("tip")=tcip
	bakrs.Update 
	else		'����Ѿ�����һ���ˣ���׷������
	bakrs("tview")=bakrs("tview")+tcid
	bakrs("tip")=bakrs("tip")+tcip
	bakrs.Update 
	end if
	bakrs.Close
	
	'����ǰ���ڵļ�¼���Ϊ�ѱ���
	conn.execute("update view set bakdays=1 where datevalue(vtime)=datevalue('" & tdate & "')")
	rs.MoveNext
next
	set tmprs=nothing


rs.Close
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
<title><%=countname%>�����ݱ��ݣ��ڶ���������ÿ������</title>
<%if bakok=false then%><meta http-equiv="refresh" content="1; url='bak_step2_in.asp?offtime=<%=offtime%>'"><%end if%>
<link rel="stylesheet" type="text/css" href="../../index.css">
</head>
<body topmargin=5 rightmargin=0 leftmargin=0 vlink=#000000 style="BACKGROUND-IMAGE: none" class=backq>
<br>
<%if bakok then%>
<p class="p1">ÿ�����ݱ�����ɡ�
<p align="right"><a href='bak_step3.asp?offtime=<%=offtime%>' target="_parent">��һ�� ���ݿͻ�����Ϣ ��ʼ</a>&nbsp;
<%else%>
<a href="bak_step2_in.asp?offtime=<%=offtime%>">ҳ��ÿ�Զ�ˢ��һ��ת��2������ݣ�ÿ��ˢ�µ�ʱ���Լ��2��30�룬������������С������ͬ������������ʱ�仹û���Զ�ˢ�£��������</a>
<%end if%>
</body>
</html>
