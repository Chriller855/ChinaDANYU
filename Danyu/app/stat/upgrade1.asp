<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!-- #include file="inc_config.asp" -->
<%

'****************** �������ݶ��� ******************

'�жϿ�ʼͳ������
	tmprs2=conn.execute("Select top 1 vtime from view order by id")
    vfirst=tmprs2("vtime")
    set tmprs2=nothing


'��ʼ����

if isupgrade() then
	on error resume next
	sql="DROP TABLE ip"
	conn.execute(sql)
end if
if isupgrade12()=false then
	on error resume next
	sql="ALTER TABLE view ADD COLUMN vwidth short"
	conn.execute(sql)
	sql="update view set vwidth=0"
	conn.execute(sql)
end if
if isupgrade13()=false then
	on error resume next
	sql="ALTER TABLE view ADD COLUMN bakdays bit"
	conn.execute(sql)
	sql="ALTER TABLE view ADD COLUMN bakstats bit"
	conn.execute(sql)
	sql="ALTER TABLE view ADD COLUMN bakpage bit"
	conn.execute(sql)
end if
if isupgrade13b()=false then
	on error resume next
	sql="ALTER TABLE vjian ADD COLUMN starttime datetime"
	conn.execute(sql)
	sql="update vjian set starttime='" & vfirst & "'"
	conn.execute(sql)
end if

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<title><%=countname%>������Ա����̨</title>
<link rel="stylesheet" type="text/css" href="../../index.css">
</head>
<body topmargin=5 rightmargin=0 leftmargin=0 vlink=#000000>
<!--#include file="inc_top.asp"-->
<br>
<table width="500" cellspacing="0" align="center" cellpadding="0" border="0">
  <tr><td colspan="3"><img src="images/photoup.gif"></td></tr>
  <tr height="30">
    <td width="1" class="backs"><img src="images/touming.gif" width="1" height="1"></td>
    <td width="498" class="backq">
		&nbsp; <img src="images/tb_index1.gif" align=absmiddle> <font style="font-size:16px">&nbsp;</font>�ӾͰ汾�������������ݿ�ṹ<br>
	<%if isupgrade() then%>
	<p class="p1">ɾ�������е�IP��ʧ�ܣ����������ݿ����ڱ���ĳ�����ã��޷�������
	<%else%>
	<p class="p1">����IP��ɾ���ɹ���
	<%end if%>
	<%if isupgrade12() then%>
	<p class="p1">����view�������ֶδ����ɹ���
	<%else%>
	<p class="p1">�������������ֶ�ʧ�ܣ����������ݿ����ڱ���ĳ�����ã��޷�������
	<%end if%>
	<%if isupgrade13() then%>
	<p class="p1">����view���ݱ���ֶδ����ɹ���
	<%else%>
	<p class="p1">�������ⱸ�ݱ���ֶ�ʧ�ܣ����������ݿ����ڱ���ĳ�����ã��޷�������
	<%end if%>
	<%if isupgrade13b() then%>
	<p class="p1">������vjian������ӡ���ʼͳ�����ڡ��ֶγɹ���
	<%else%>
	<p class="p1">������vjian������ӡ���ʼͳ�����ڡ��ֶ�ʧ�ܣ����������ݿ����ڱ���ĳ�����ã��޷�������
	<%end if%>
	</td>
    <td width="1" class="backs"><img src="images/touming.gif" width="1" height="1"></td>
  </tr>
  <tr><td colspan="4"><img src="images/photodown.gif"></td></tr>
</table>
<br>
<!--#include file="inc_bottom.asp"-->
</body>
</html>
<%
'****************** �ر����ݶ��� ******************
conn.Close
set conn=nothing

'****************** �Զ��庯�� ********************
function isUpgrade()
	isUpgrade=true
	on error resume next
	sql="select * from ip"
	conn.execute(sql)
	if err<>0 then isUpgrade=false
end function

function isUpgrade12()
	isUpgrade12=true
	on error resume next
	sql="select vwidth from view"
	conn.execute(sql)
	if err<>0 then isUpgrade12=false
end function

function isUpgrade13()
	isUpgrade13=true
	on error resume next
	sql="select bakdays from view"
	conn.execute(sql)
	if err<>0 then isUpgrade13=false
end function

function isUpgrade13b()
	isUpgrade13b=true
	on error resume next
	sql="select starttime from vjian"
	conn.execute(sql)
	if err<>0 then isUpgrade13b=false
end function

%>