<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!-- #include file="inc_config.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<title><%=countname%>�������ߵ���ͳ��</title>
<link rel="stylesheet" type="text/css" href="../../index.css">
</head>
<body topmargin=5 rightmargin=0 leftmargin=0 vlink=#000000>
<%
Set rs = Server.CreateObject("ADODB.Recordset")
%>

<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr height="30">
    <td>�����ߵ�����������</td>
  </tr>
  <tr height="30"> 
    <td><a href="tj_where.asp">������������</a> | <a href="tj_where.asp?iSort=1">����������</a> 
      | <a href="tj_where.asp?iSort=2">�Ϻ�����</a></td>
  </tr>
  <tr height="30"> 
    <td width="498"> 
      <table border="0" cellpadding="0" cellspacing="0" width="350">
        <tr height="9"> 
          <td></td>
        </tr>
        <tr height="10"> 
          <td width="120"></td>
          <td width="230"><img src="images/tu_back_up.gif"></td>
        </tr>
        <%
if request("iSort")="" then	sql="select vwhere,count(id) as allwhere from view group by vwhere order by count(id) DESC"
if request("iSort")="1" then sql="select vwhere,count(id) as allwhere from view group by vwhere order by vwhere asc, count(id) DESC"
if request("iSort")="2" then sql="select vwhere,count(id) as allwhere from view where vwhere like '�Ϻ�%' group by vwhere order by count(id) DESC"

rs.Open sql,conn,1,1
maxallwhere=0
sumallwhere=0
do while not rs.EOF
	if clng(rs("allwhere"))>maxallwhere then maxallwhere=clng(rs("allwhere"))
	sumallwhere=sumallwhere+clng(rs("allwhere"))
	rs.MoveNext
loop
	'��ֹ����Ϊ0������
	if maxallwhere=0 then maxallwhere=1
	if sumallwhere=0 then sumallwhere=1
rs.MoveFirst 

j=0
do while not rs.EOF
thewhere=rs("vwhere")
vallwhere=rs("allwhere")
	thelen=len(thewhere)
	if thelen =0 then
		thewhere="main.asp"
		svwhere="ͨ���ղػ�ֱ��������ַ����"
	end if
	if thelen <= 33 and thelen > 0 then
		svwhere=thewhere
	end if
	if thelen >= 34 then
		svwhere=left(thewhere,31) & "..."
	end if
%>
        <tr> 
          <td width="120" align=right><a title="<%=thewhere%>������<%=vallwhere%>�Σ�<%
	'����������İٷ�������ȷ��С����1λ��С�������ǰ�����ĸ0
	lsbf=int(vallwhere*1000/sumallwhere)/10
	if lsbf<1 then lsbf="0" & lsbf
	Response.Write lsbf
	%>%"><%=svwhere%></a>&nbsp;</td>
          <td width="230" background="images/tu_back_2.gif" align=left><img style="BORDER-left: #000000 1px solid;" src="images/tu.gif"
	width="<%=(vallwhere/maxallwhere)*183%>" height="9"
	alt="<%=thewhere%>������<%=vallwhere%>�Σ�<%
	'����������İٷ�������ȷ��С����1λ��С�������ǰ�����ĸ0
	lsbf=int(vallwhere*1000/sumallwhere)/10
	if lsbf<1 then lsbf="0" & lsbf
	Response.Write lsbf
	%>%"> <%=vallwhere%></td>
        </tr>
        <%
rs.MoveNext
	j=j+1
	'�����¼����40�������˳�
	if j=200 then exit do
loop
%>
        <tr height="10"> 
          <td width="120"></td>
          <td width="230"><img src="images/tu_back_down.gif"></td>
        </tr>
        <tr height="5"> 
          <td colspan=29></td>
        </tr>
      </table></td>
  </tr>
</table>
������ָͼ�����������ֿ��Կ�����Ӧ�ķ�������Ҫ�õ�ĳһʱ�ε�ͳ����Ϣ����ʹ���Զ���ͳ�ơ�
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