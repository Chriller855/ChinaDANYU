<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!-- #include file="inc_config.asp" -->
<%
'Ȩ�޼��
if session.Contents("master")=false and whatcan<2 then Response.Redirect "help.asp?id=004&error=��û�в鿴��·ͳ�Ƶ�Ȩ�ޡ�"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<title><%=countname%>����·������ͳ��</title>
<link rel="stylesheet" type="text/css" href="../../index.css">
</head>
<body topmargin=5 rightmargin=0 leftmargin=0 vlink=#000000>
<%
Set rs = Server.CreateObject("ADODB.Recordset")
%>

<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr height="30">
    <td width="498">��·��������<br>
        <table border="0" cellpadding="0" cellspacing="0" width="450">
          <tr height="9">
            <td></td>
          </tr>
          <tr height="10">
            <td></td>
            <td width="230"><img src="images/tu_back_up.gif"></td>
          </tr>
          <%

sql="select vcome,count(id) as allcome from view group by vcome order by count(id) DESC"
rs.Open sql,conn,1,1

maxallcome=0
sumallcome=0
do while not rs.EOF
	if clng(rs("allcome"))>maxallcome then maxallcome=clng(rs("allcome"))
	sumallcome=sumallcome+clng(rs("allcome"))
	rs.MoveNext
loop
	'��ֹ����Ϊ0������
	if maxallcome=0 then maxallcome=1
	if sumallcome=0 then sumallcome=1
rs.MoveFirst 

j=0
do while not rs.EOF
thecome=rs("vcome")
vallcome=rs("allcome")
	thelen=len(thecome)
	if thelen =0 then
		thecome="main.asp"
		svcome="ͨ���ղػ�ֱ��������ַ����"
	end if
	if thelen <= 33 and thelen > 0 then
		svcome=thecome
	end if
	if thelen >= 34 then
		svcome=left(thecome,31) & "..."
	end if
%>
          <tr>
            <td align=right><a href="<%=thecome%>" target="_blank"
	title="<%=thecome & vbcrlf%>����<%=vallcome%>�Σ�<%
	'����������İٷ�������ȷ��С����1λ��С�������ǰ�����ĸ0
	lsbf=int(vallcome*1000/sumallcome)/10
	if lsbf<1 then lsbf="0" & lsbf
	Response.Write lsbf
	%>%"><%=svcome%></a>&nbsp;</td>
            <td width="230" background="images/tu_back_2.gif" align=left><img style="BORDER-left: #000000 1px solid;" src="images/tu.gif"
	width="<%=(vallcome/maxallcome)*183%>" height="9"
	alt="<%=thecome & vbcrlf%>����<%=vallcome%>�Σ�<%
	'����������İٷ�������ȷ��С����1λ��С�������ǰ�����ĸ0
	lsbf=int(vallcome*1000/sumallcome)/10
	if lsbf<1 then lsbf="0" & lsbf
	Response.Write lsbf
	%>%"> <%=vallcome%></td>
          </tr>
          <%
rs.MoveNext
	j=j+1
	'�����¼����40�������˳�
	if j=40 then exit do
loop
%>
          <tr height="10">
            <td></td>
            <td width="230"><img src="images/tu_back_down.gif"></td>
          </tr>
          <tr height="5">
            <td colspan=29></td>
          </tr>
      </table></td>
  </tr>
</table>
������ָͼ����������ַ���Կ�����Ӧ�ķ�������Ҫ�õ�ĳһʱ�ε�ͳ����Ϣ����ʹ���Զ���ͳ�ơ�
<%
rs.Close

set rs=nothing
conn.Close 
set conn=nothing
%>
<br>
</body>
</html>
<%
'����ָ�����ڵķ�����
function vdaycon(theday)
    theday=cdate(theday)
    thetday=cdate(theday+1)
    tmprs=conn.execute("Select count(id) as vdaycon from view where" & _
		" vtime>=datevalue('" & theday & "') and vtime<=datevalue('" & thetday & "')")
    vdaycon=tmprs("vdaycon")
	if isnull(vdaycon) then vdaycon=0
end function
%>