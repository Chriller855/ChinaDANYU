<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!-- #include file="inc_config.asp" -->
<%
'Ȩ�޼��
if session.Contents("master")=false and whatcan<2 then Response.Redirect "help.asp?id=004&error=��û�в鿴��ͳ�Ƶ�Ȩ�ޡ�"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<title><%=countname%>���·���ͳ��</title>
<link rel="stylesheet" type="text/css" href="../../index.css">
</head>
<body topmargin=5 rightmargin=0 leftmargin=0 vlink=#000000>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr height="30">
    <td>���12���·�����<br>
        <table border="0" cellpadding="0" cellspacing="0" width="310">
          <tr height="9">
            <td></td>
          </tr>
          <tr height="101">
            <%
Set rs = Server.CreateObject("ADODB.Recordset")
'����12���£���ͷҲ��һ���£�ǰ��ʱ��
datetwelve=dateadd("m",-11,date())
datetwelve=cdate(year(datetwelve) & "-" & month(datetwelve) & "-1")

sql="select vmonth,count(id) as allmonth from view where vtime>=datevalue('" & _
	datetwelve & "') group by vmonth"
rs.Open sql,conn,1,1

dim vallmonth(12)
maxallmonth=0
sumallmonth=0
do while not rs.EOF
	vallmonth(clng(rs("vmonth"))-1)=clng(rs("allmonth"))
	if vallmonth(clng(rs("vmonth"))-1)>maxallmonth then maxallmonth=vallmonth(clng(rs("vmonth"))-1)
	sumallmonth=sumallmonth+vallmonth(clng(rs("vmonth"))-1)
	rs.MoveNext
loop
	'��ֹ����Ϊ�������
	if maxallmonth=0 then maxallmonth=1
	if sumallmonth=0 then sumallmonth=1

%>
            <td align=right valign=top><p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 13"> <font face="Arial"><%=int(maxallmonth*10+0.5)/10%></font>
                <p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 13"> <font face="Arial"><%=int(3*maxallmonth*10/4+0.5)/10%></font>
                <p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 13"> <font face="Arial"><%=int(maxallmonth*10/2+0.5)/10%></font>
                <p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 0"> <font face="Arial"><%=int(maxallmonth*10/4+0.5)/10%><br>
              </font></td>
            <td width=10><img src="images/tu_back_left.gif"></td>
            <%
for i= 0 to 11
themonth=month(now())+i
if themonth>11 then themonth=themonth-12

%>
            <td width=20 valign=bottom background="images/tu_back.gif" align=center><img style="BORDER-BOTTOM: #000000 1px solid;" src="images/tu.gif"
	height="<%=(vallmonth(themonth)/maxallmonth)*100%>" width="9"
	alt="<%=themonth+1%>�£�����<%=vallmonth(themonth)%>�Σ�<%
	'����������İٷ�������ȷ��С����1λ��С�������ǰ�����ĸ0
	lsbf=int(vallmonth(themonth)*1000/sumallmonth)/10
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
for i= 0 to 11
themonth=month(now())+i
if themonth>11 then themonth=themonth-12
%>
            <td width=20 align=center><a title="<%=themonth+1%>�£�����<%=vallmonth(themonth)%>�Σ�<%
	'����������İٷ�������ȷ��С����1λ��С�������ǰ�����ĸ0
	lsbf=int(vallmonth(themonth)*1000/sumallmonth)/10
	if lsbf<1 then lsbf="0" & lsbf
	Response.Write lsbf
	%>%"> <font face="Arial" style="letter-spacing: -1"><%=themonth+1%></font></a></td>
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
<%
rs.Close
%>


<br>



<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr height="30">
    <td>����12���·�����<br>
        <table border="0" cellpadding="0" cellspacing="0" width="310">
          <tr height="9">
            <td></td>
          </tr>
          <tr height="101">
            <%

sql="select vmonth,count(id) as allmonth from view group by vmonth"
rs.Open sql,conn,1,1

for i=0 to 11
	vallmonth(i)=0
next
maxallmonth=0
sumallmonth=0
do while not rs.EOF
	vallmonth(clng(rs("vmonth"))-1)=clng(rs("allmonth"))
	if vallmonth(clng(rs("vmonth"))-1)>maxallmonth then maxallmonth=vallmonth(clng(rs("vmonth"))-1)
	sumallmonth=sumallmonth+vallmonth(clng(rs("vmonth"))-1)
	rs.MoveNext
loop
	'��ֹ����Ϊ�������
	if maxallmonth=0 then maxallmonth=1
	if sumallmonth=0 then sumallmonth=1

%>
            <td align=right valign=top><p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 13"> <font face="Arial"><%=int(maxallmonth*10+0.5)/10%></font>
                <p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 13"> <font face="Arial"><%=int(3*maxallmonth*10/4+0.5)/10%></font>
                <p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 13"> <font face="Arial"><%=int(maxallmonth*10/2+0.5)/10%></font>
                <p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 0"> <font face="Arial"><%=int(maxallmonth*10/4+0.5)/10%><br>
              </font></td>
            <td width=10><img src="images/tu_back_left.gif"></td>
            <%
for i= 0 to 11
themonth=i

%>
            <td width=20 valign=bottom background="images/tu_back.gif" align=center><img style="BORDER-BOTTOM: #000000 1px solid;" src="images/tu.gif"
	height="<%=(vallmonth(themonth)/maxallmonth)*100%>" width="9"
	alt="<%=themonth+1%>�£�����<%=vallmonth(themonth)%>�Σ�<%
	'����������İٷ�������ȷ��С����1λ��С�������ǰ�����ĸ0
	lsbf=int(vallmonth(themonth)*1000/sumallmonth)/10
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
for i= 0 to 11
themonth=i
%>
            <td width=20 align=center><a title="<%=themonth+1%>�£�����<%=vallmonth(themonth)%>�Σ�<%
	'����������İٷ�������ȷ��С����1λ��С�������ǰ�����ĸ0
	lsbf=int(vallmonth(themonth)*1000/sumallmonth)/10
	if lsbf<1 then lsbf="0" & lsbf
	Response.Write lsbf
	%>%"> <font face="Arial" style="letter-spacing: -1"><%=themonth+1%></font></a></td>
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
<%
rs.Close
%>

<br>

<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr height="30">
    <td>�������ͳ��<br>
        <table border="0" cellpadding="0" cellspacing="0" width="270">
          <tr height="9">
            <td></td>
          </tr>
          <tr height="10">
            <td width="40"></td>
            <td width="230"><img src="images/tu_back_up.gif"></td>
          </tr>
          <%

sql="select vyear,count(id) as allyear from view group by vyear order by vyear DESC"
rs.Open sql,conn,1,1

maxallyear=0
sumallyear=0
do while not rs.EOF
	if clng(rs("allyear"))>maxallyear then maxallyear=clng(rs("allyear"))
	sumallyear=sumallyear+clng(rs("allyear"))
	rs.MoveNext
loop
	'��ֹ����Ϊ�������
	if maxallyear=0 then maxallyear=1
	if sumallyear=0 then sumallyear=1

rs.MoveFirst 
do while not rs.EOF
theyear=rs("vyear")
vallyear=rs("allyear")
%>
          <tr>
            <td width="40" align=right><a title="<%=theyear%>�꣬����<%=vallyear%>�Σ�<%
	'����������İٷ�������ȷ��С����1λ��С�������ǰ�����ĸ0
	lsbf=int(vallyear*1000/sumallyear)/10
	if lsbf<1 then lsbf="0" & lsbf
	Response.Write lsbf
	%>%"><%=theyear%></a>&nbsp;</td>
            <td width="230" background="images/tu_back_2.gif" align=left><img style="BORDER-left: #000000 1px solid;" src="images/tu.gif"
	width="<%=(vallyear/maxallyear)*183%>" height="9"
	alt="<%=theyear%>�꣬����<%=vallyear%>�Σ�<%
	'����������İٷ�������ȷ��С����1λ��С�������ǰ�����ĸ0
	lsbf=int(vallyear*1000/sumallyear)/10
	if lsbf<1 then lsbf="0" & lsbf
	Response.Write lsbf
	%>%"> <%=vallyear%></td>
          </tr>
          <%
rs.MoveNext
loop
%>
          <tr height="10">
            <td width="40"></td>
            <td width="230"><img src="images/tu_back_down.gif"></td>
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