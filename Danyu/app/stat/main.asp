<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!-- #include file="inc_config.asp" -->
<%
'Ȩ�޷���
if session.Contents("master")=false and whatcan<1 then Response.Redirect "help.asp?id=004&error=��ͳ��ϵͳ����Ա������ÿͲ鿴�κ���Ϣ��"

'�����ܷ���������ʼ�������ڣ��������ݿ��ȡ��
    tmprs=conn.execute("Select vtop,starttime from vjian")
    vtotal=tmprs("vtop")
    vfirst=tmprs("starttime")
	set tmprs=nothing
	if isnull(vtotal) then vtotal=0

if vtotal=0 then
	conn.Close
	set conn=nothing
	Response.Redirect "help.asp?id=007&error=ͳ��ϵͳ��û�����ã��в��ܲ鿴ͳ�Ʊ��档"
end if


'��������
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql="select vip from view where vtime >= dateadd('n',-20,now()) group by vip"
	rs.Open sql,conn,1,1
	vonline=rs.RecordCount
	rs.Close
	set rs=nothing
'���շ����������շ��������Ӽ����ݳ���ȡ��
    tmprs=conn.execute("Select today,yesterday from vjian")
    vtoday=tmprs("today")
    vyesterday=tmprs("yesterday")
	if isnull(vtoday) then vtoday=0
	if isnull(vyesterday) then vyesterday=0
'���������
    tmprs=conn.execute("Select count(id) as vthisyear from view where vyear=" & year(now))
    vthisyear=tmprs("vthisyear")
	if isnull(vthisyear) then vthisyear=0
'���·�����
    tmprs=conn.execute("Select count(id) as vthismonth from view where vmonth=" & month(now) & " and vyear=" & year(now))
    vthismonth=tmprs("vthismonth")
	if isnull(vthismonth) then vthismonth=0
'����������ƽ��ÿ�������
	vdays=now()-vfirst
	vdayavg=vtotal/vdays
	vdays=int((vdays*10^mPrecision)+0.5)/(10^mPrecision)
	if vdays<1 then vdays="0" & vdays
	vdayavg=int((vdayavg*10^mPrecision)+0.5)/(10^mPrecision)
'Ԥ�ƽ��շ�����
	vdaylong=now()-date()
	vguess=int(((vtoday/vdaylong)+vyesterday)/2+0.5)
	if vguess< vtoday then vguess=int((vtoday/vdaylong)+0.5)
'��ǰ�û�����
	vuser=cint(Request.Cookies(mNameEn)("lao"))

conn.Close
set conn=nothing
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<title><%=countname%></title>
<link rel="stylesheet" type="text/css" href="../../index.css">
</head>
<body topmargin=5 rightmargin=0 leftmargin=0 vlink=#000000>
<table width="98%"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td bgcolor="#CCCCCC">�������� </td>
  </tr>
  <tr>
    <td>�ܷ�����: &nbsp; <%=vtotal+old_count%><br>
�ñ�ϵͳǰ: <%=old_count%><br>
�ñ�ϵͳ��: <%=vtotal%><br>
��������: &nbsp; <%=vonline%><br>
���ķ�����: <%=vuser%><br>
��ʼͳ����: <%=vfirst%><br>
���շ�����: <%=vtoday%><br>
���շ�����: <%=vyesterday%><br>
���������: <%=vthisyear%><br>
���·�����: <%=vthismonth%><br>
ͳ������: &nbsp; <%=vdays%><br>
ƽ���շ���: <%=vdayavg%><br>
Ԥ�ƽ���: &nbsp; <%=vguess%></td>
  </tr>
</table>
</body>
</html>
