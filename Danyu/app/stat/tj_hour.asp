<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!-- #include file="inc_config.asp" -->
<%
'Ȩ�޼��
if session.Contents("master")=false and whatcan<2 then Response.Redirect "help.asp?id=004&error=��û�в鿴24Сʱ����ͳ�Ƶ�Ȩ�ޡ�"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<title><%=countname%>�����/����24Сʱ����ͳ��</title>
<link rel="stylesheet" type="text/css" href="../../index.css">
</head>
<body topmargin=5 rightmargin=0 leftmargin=0 vlink=#000000>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr height="30">
    <td width="498">���24Сʱ������<br>
        <table border="0" cellpadding="0" cellspacing="0" width="430">
          <tr height="9">
            <td colspan=29></td>
          </tr>
          <tr height="101">
            <%
dim vhour(24)
maxhour=0
sumhour=0
for i=0 to 23
	vhour(i)=vhourcon(i)
	if vhour(i)>maxhour then maxhour=vhour(i)
	sumhour=sumhour+vhour(i)
next
'��ֹ����Ϊ0����
if maxhour=0 then maxhour=1
if sumhour=0 then sumhour=1

%>
            <td align=right valign=top><p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 13"> <font face="Arial"><%=int(maxhour*10+0.5)/10%></font>
                <p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 13"> <font face="Arial"><%=int(3*maxhour*10/4+0.5)/10%></font>
                <p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 13"> <font face="Arial"><%=int(maxhour*10/2+0.5)/10%></font>
                <p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 0"> <font face="Arial"><%=int(maxhour*10/4+0.5)/10%><br>
              </font></td>
            <td width=10><img src="images/tu_back_left.gif"></td>
            <%
for i= 0 to 23
thehour=hour(now())+i+1
if thehour>23 then thehour=thehour-24
%>
            <td width=15 valign=bottom background="images/tu_back.gif" align=center><img style="BORDER-BOTTOM: #000000 1px solid;" src="images/tu.gif"
	height="<%=(vhour(thehour)/maxhour)*100%>" width="9"
	alt="<%=thehour%>ʱ������<%=vhour(thehour)%>�Σ�<%
	'����������İٷ�������ȷ��С����1λ��С�������ǰ�����ĸ0
	lsbf=int(vhour(thehour)*1000/sumhour)/10
	if lsbf<1 then lsbf="0" & lsbf
	Response.Write lsbf
	%>%"></td>
            <%
next
%>
            <td width=10><img src="images/tu_back_right.gif"></td>
            <td width=10></td>
          </tr>
          <tr>
            <td align=right><p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 0"> <font face="Arial">0</font></td>
            <td width=10></td>
            <%
for i= 0 to 23
thehour=hour(now())+i+1
if thehour>23 then thehour=thehour-24
%>
            <td width=15 align=center><a title="<%=thehour%>ʱ������<%=vhour(thehour)%>�Σ�<%
	'����������İٷ�������ȷ��С����1λ��С�������ǰ�����ĸ0
	lsbf=int(vhour(thehour)*1000/sumhour)/10
	if lsbf<1 then lsbf="0" & lsbf
	Response.Write lsbf
	%>%"><font face="Arial" style="letter-spacing: -1"><%=thehour%></font></a></td>
            <%
next
%>
            <td width=10></td>
            <td width=10></td>
          </tr>
          <tr height="5">
            <td colspan=29></td>
          </tr>
      </table></td>
  </tr>
</table>
<br>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr height="30">
    <td width="498"> ����24Сʱ������<br>
        <table border="0" cellpadding="0" cellspacing="0" width="430">
          <tr height="9">
            <td colspan=29></td>
          </tr>
          <tr height="101">
            <%
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select vhour,count(id) as allhour from view group by vhour"
rs.Open sql,conn,1,1

dim vallhour(24)
maxallhour=0
sumallhour=0
do while not rs.EOF
	vallhour(clng(rs("vhour")))=clng(rs("allhour"))
	if vallhour(clng(rs("vhour")))>maxallhour then maxallhour=vallhour(clng(rs("vhour")))
	sumallhour=sumallhour+vallhour(clng(rs("vhour")))
	rs.MoveNext
loop
'��ֹ����Ϊ0����
if maxallhour=0 then maxallhour=1
if sumallhour=0 then sumallhour=1
%>
            <td align=right valign=top><p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 13"> <font face="Arial"><%=int(maxallhour*10+0.5)/10%></font>
                <p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 13"> <font face="Arial"><%=int(3*maxallhour*10/4+0.5)/10%></font>
                <p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 13"> <font face="Arial"><%=int(maxallhour*10/2+0.5)/10%></font>
                <p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 0"> <font face="Arial"><%=int(maxallhour*10/4+0.5)/10%><br>
              </font></td>
            <td width=10><img src="images/tu_back_left.gif"></td>
            <%
for i= 0 to 23
%>
            <td width=15 valign=bottom background="images/tu_back.gif" align=center><img style="BORDER-BOTTOM: #000000 1px solid;" src="images/tu.gif"
	height="<%=(vallhour(i)/maxallhour)*100%>" width="9"
	alt="<%=i%>ʱ������<%=vallhour(i)%>�Σ�<%
	'����������İٷ�������ȷ��С����1λ��С�������ǰ�����ĸ0
	lsbf=int(vallhour(i)*1000/sumallhour)/10
	if lsbf<1 then lsbf="0" & lsbf
	Response.Write lsbf
	%>%"></td>
            <%
next
%>
            <td width=10><img src="images/tu_back_right.gif"></td>
            <td width=10></td>
          </tr>
          <tr>
            <td align=right><p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 0"> <font face="Arial">0</font></td>
            <td width=10></td>
            <%
for i= 0 to 23
%>
            <td width=15 align=center><a title="<%=i%>ʱ������<%=vallhour(i)%>�Σ�<%
	'����������İٷ�������ȷ��С����1λ��С�������ǰ�����ĸ0
	lsbf=int(vallhour(i)*1000/sumallhour)/10
	if lsbf<1 then lsbf="0" & lsbf
	Response.Write lsbf
	%>%"><font face="Arial" style="letter-spacing: -1"><%=i%></font></a></td>
            <%
next
%>
            <td width=10></td>
            <td width=10></td>
          </tr>
          <tr height="5">
            <td colspan=29></td>
          </tr>
      </table></td>
  </tr>
</table>
������ָͼ��������ͼ�����µ����ֿ��Կ�����Ӧ�ķ�������

<%
rs.Close
set rs=nothing
conn.Close 
set conn=nothing
%>
</body>
</html>
<%
function vhourcon(thehour)
  if thehour=hour(now()) then
    tmprs=conn.execute("Select count(id) as vhourcon from view where vhour=" & _
		hour(now()) & " and vday=" & day(now()) & " and vmonth=" & month(now()) & " and vyear=" & year(now()))
    vhourcon=tmprs("vhourcon")
	if isnull(vhourcon) then vhourcon=0
  else
    tmprs=conn.execute("Select count(id) as vhourcon from view where vhour=" & _
		thehour & " and vtime>now()-1")
    vhourcon=tmprs("vhourcon")
	if isnull(vhourcon) then vhourcon=0
  end if
end function
%>