<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!-- #include file="inc_config.asp" -->
<%
'Ȩ�޼��
if session.Contents("master")=false and whatcan<2 then Response.Redirect "help.asp?id=004&error=��û�в鿴�ͻ������ͳ�Ƶ�Ȩ�ޡ�"
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<title><%=countname%>���ͻ������������ͳ��</title>
<link rel="stylesheet" type="text/css" href="../../index.css">
</head>
<body topmargin=5 rightmargin=0 leftmargin=0 vlink=#000000>
<%'�����ʹ�����

dim soft(7,2)

	soft(0,0)="NetCaptor"
	soft(1,0)="MSIE 6.x"
	soft(2,0)="MSIE 5.x"
	soft(3,0)="MSIE 4.x"
	soft(4,0)="Netscape"
	soft(5,0)="Opera"
	soft(6,0)="Other"

	'����������д�����飬howsoft
	soft(0,1)=howsoft("NetCaptor")
	soft(1,1)=howsoft("MSIE 6.x")
	soft(2,1)=howsoft("MSIE 5.x")
	soft(3,1)=howsoft("MSIE 4.x")
	soft(4,1)=howsoft("Netscape")
	soft(5,1)=howsoft("Opera")
	soft(6,1)=howsoft("Other")

	'�ҵ����ֵ����ֵ
	maxsoft=0
	sumsoft=0
	for i=0 to 6
		if soft(i,1)>maxsoft then maxsoft=soft(i,1)
		sumsoft=sumsoft+soft(i,1)
	next
	'��ֹ����Ϊ0������
	if maxsoft=0 then maxsoft=1
	if sumsoft=0 then sumsoft=1
%>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr height="30">
    <td width="498">�������������<br>
        <table border="0" cellpadding="0" cellspacing="0" width="385">
          <tr height="9">
            <td></td>
          </tr>
          <tr height="101">
            <td align=right valign=top><p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 13"> <font face="Arial"><%=int(maxsoft*10+0.5)/10%></font>
                <p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 13"> <font face="Arial"><%=int(3*maxsoft*10/4+0.5)/10%></font>
                <p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 13"> <font face="Arial"><%=int(maxsoft*10/2+0.5)/10%></font>
                <p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 0"> <font face="Arial"><%=int(maxsoft*10/4+0.5)/10%><br>
              </font></td>
            <td width=10><img src="images/tu_back_left.gif"></td>
            <%
for i= 0 to 6
%>
            <td width=45 valign=bottom background="images/tu_back.gif" align=center><img style="BORDER-BOTTOM: #000000 1px solid" src="images/tu.gif"
	height="<%=(soft(i,1)/maxsoft)*100%>" width="9"
	alt="<%=soft(i,0)%>������<%=soft(i,1)%>�Σ�<%
	
	'����������İٷ�������ȷ��С����1λ��С�������ǰ�����ĸ0
	lsbf=int(soft(i,1)*1000/sumsoft+0.5)/10
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
for i= 0 to 6
%>
            <td width=45 align=center><a title="<%=soft(i,0)%>������<%=soft(i,1)%>�Σ�<%
	
	'����������İٷ�������ȷ��С����1λ��С�������ǰ�����ĸ0
	lsbf=int(soft(i,1)*1000/sumsoft+0.5)/10
	if lsbf<1 then lsbf="0" & lsbf
	Response.Write lsbf
	
	%>%"><%=left(soft(i,0),6)%></a></td>
            <%next%>
            <td width=10></td>
            <td width=10></td>
          </tr>
          <tr height="5">
            <td></td>
          </tr>
      </table></td>
  </tr>
</table>
<br>



<%'����ϵͳʹ�����

dim OS(7,2)

	OS(0,0)="Win 2000"
	OS(1,0)="Win XP"
	OS(2,0)="Win NT"
	OS(3,0)="Win 9x"
	OS(4,0)="��Unix"
	OS(5,0)="Mac"
	OS(6,0)="Other"

	'������ϵͳ���д�����飬howOS
	OS(0,1)=howOS("Win 2000")
	OS(1,1)=howOS("Win XP")
	OS(2,1)=howOS("Win NT")
	OS(3,1)=howOS("Win 9x")
	OS(4,1)=howOS("��Unix")
	OS(5,1)=howOS("Mac")
	OS(6,1)=howOS("Other")

	'�ҵ����ֵ����ֵ
	maxOS=0
	sumOS=0
	for i=0 to 6
		if OS(i,1)>maxOS then maxOS=OS(i,1)
		sumOS=sumOS+OS(i,1)
	next
	'��ֹ����Ϊ0������
	if maxOS=0 then maxOS=1
	if sumOS=0 then sumOS=1
%>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr height="30">
    <td width="498">����ϵͳ��������<br>
        <table border="0" cellpadding="0" cellspacing="0" width="385">
          <tr height="9">
            <td></td>
          </tr>
          <tr height="101">
            <td align=right valign=top><p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 13"> <font face="Arial"><%=int(maxOS*10+0.5)/10%></font>
                <p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 13"> <font face="Arial"><%=int(3*maxOS*10/4+0.5)/10%></font>
                <p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 13"> <font face="Arial"><%=int(maxOS*10/2+0.5)/10%></font>
                <p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 0"> <font face="Arial"><%=int(maxOS*10/4+0.5)/10%><br>
              </font></td>
            <td width=10><img src="images/tu_back_left.gif"></td>
            <%
for i= 0 to 6
%>
            <td width=45 valign=bottom background="images/tu_back.gif" align=center><img style="BORDER-BOTTOM: #000000 1px solid" src="images/tu.gif"
	height="<%=(OS(i,1)/maxOS)*100%>" width="9"
	alt="<%=OS(i,0)%>������<%=OS(i,1)%>�Σ�<%
	
	'����������İٷ�������ȷ��С����1λ��С�������ǰ�����ĸ0
	lsbf=int(OS(i,1)*1000/sumOS+0.5)/10
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
for i= 0 to 6
%>
            <td width=45 align=center><a title="<%=OS(i,0)%>������<%=OS(i,1)%>�Σ�<%
	
	'����������İٷ�������ȷ��С����1λ��С�������ǰ�����ĸ0
	lsbf=int(OS(i,1)*1000/sumOS+0.5)/10
	if lsbf<1 then lsbf="0" & lsbf
	Response.Write lsbf
	
	%>%">
              <%if i=0 then%>
              Win 2k
              <%else%>
              <%=OS(i,0)%>
              <%end if%>
            </a></td>
            <%next%>
            <td width=10></td>
            <td width=10></td>
          </tr>
          <tr height="5">
            <td></td>
          </tr>
      </table></td>
  </tr>
</table>
<br>

<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr height="30">
    <td width="498">�ͻ�����Ļ���ͳ��<br>
        <table border="0" cellpadding="0" cellspacing="0" width="270">
          <tr height="9">
            <td></td>
          </tr>
          <tr height="10">
            <td width="40"></td>
            <td width="230"><img src="images/tu_back_up.gif"></td>
          </tr>
          <%


Set rs = Server.CreateObject("ADODB.Recordset")
sql="select vwidth,count(id) as allwidth from view where vwidth<>0 group by vwidth order by vwidth DESC"
rs.Open sql,conn,1,1

maxallwidth=0
sumallwidth=0
do while not rs.EOF
	if clng(rs("allwidth"))>maxallwidth then maxallwidth=clng(rs("allwidth"))
	sumallwidth=sumallwidth+clng(rs("allwidth"))
	rs.MoveNext
loop
	'��ֹ����Ϊ�������
	if maxallwidth=0 then maxallwidth=1
	if sumallwidth=0 then sumallwidth=1

rs.MoveFirst 
do while not rs.EOF
thewidth=rs("vwidth")
vallwidth=rs("allwidth")
%>
          <tr>
            <td width="40" align=right><a title="<%=thewidth%>���أ�<%=vallwidth%>�Σ�<%
	'����������İٷ�������ȷ��С����1λ��С�������ǰ�����ĸ0
	lsbf=int(vallwidth*1000/sumallwidth)/10
	if lsbf<1 then lsbf="0" & lsbf
	Response.Write lsbf
	%>%"><%=thewidth%></a>&nbsp;</td>
            <td width="230" background="images/tu_back_2.gif" align=left><img style="BORDER-left: #000000 1px solid;" src="images/tu.gif"
	width="<%=(vallwidth/maxallwidth)*183%>" height="9"
	alt="<%=thewidth%>���أ�<%=vallwidth%>�Σ�<%
	'����������İٷ�������ȷ��С����1λ��С�������ǰ�����ĸ0
	lsbf=int(vallwidth*1000/sumallwidth)/10
	if lsbf<1 then lsbf="0" & lsbf
	Response.Write lsbf
	%>%"> <%=vallwidth%></td>
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
������ָͼ�������߶�Ӧ�����ֿ��Կ�����Ӧ�ķ�������
<%
rs.Close
set rs=nothing
conn.Close 
set conn=nothing
%>
</body>
</html>
<%
'����ʹ��ָ��������������
function howsoft(vsoft)
    tmprs=conn.execute("Select count(id) as howsoft from view where" & _
		" vsoft='" & vsoft & "'")
    howsoft=tmprs("howsoft")
	if isnull(howsoft) then howsoft=0
end function

'����ʹ��ָ������ϵͳ�������
function howOS(vOS)
    tmprs=conn.execute("Select count(id) as howOS from view where" & _
		" vOS='" & vOS & "'")
    howOS=tmprs("howOS")
	if isnull(howOS) then howOS=0
end function
%>