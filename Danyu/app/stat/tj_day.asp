<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!-- #include file="inc_config.asp" -->
<%
'Ȩ�޼��
if session.Contents("master")=false and whatcan<2 then Response.Redirect "help.asp?id=004&error=��û�в鿴�շ���ͳ�Ƶ�Ȩ�ޡ�"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<title><%=countname%>���շ���ͳ��</title>
<link rel="stylesheet" type="text/css" href="../../index.css">
</head>
<body topmargin=5 rightmargin=0 leftmargin=0 vlink=#000000>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr height="30">
    <td width="498">���31�������<br>
        <table border="0" cellpadding="0" cellspacing="0" width="453">
          <tr height="9">
            <td></td>
          </tr>
          <tr height="101">
            <%
'�ҵ���ʼͳ�������������������31�죬������ǰ��Ŀռ�
tmprs=conn.execute("Select top 1 vtime from view order by id")
vfirst=tmprs("vtime")
set tmprs=nothing

if isnull(vfirst) then vfirst=now()
vdays=int(date()-vfirst+1)

'������ά���飬voutday(*,0)Ϊ������,voutday(*,1)Ϊ����,voutday(*,2)Ϊ����
dim vday(31,3),voutday(31,3)
maxday=0
sumday=0
for i=0 to 30
	vday(i,0)=vdaycon(date()-30+i)
	if vday(i,0)>maxday then maxday=vday(i,0)
	sumday=sumday+vday(i,0)
	vday(i,1)=day(date()-30+i)
	vday(i,2)=weekday(date()-30+i)
next
	'��ֹ����Ϊ0������
	if maxday=0 then maxday=1
	if sumday=0 then sumday=1

'������ͳ����������ֵ����
if vdays>=31 then
	for i=0 to 30
		voutday(i,0)=vday(i,0)
		voutday(i,1)=vday(i,1)
		voutday(i,2)=vday(i,2)
	next
else
	on error resume next
	for i=0 to 30
		if i<=vdays then
			voutday(i,0)=vday(i+30-vdays,0)
			voutday(i,1)=vday(i+30-vdays,1)
			voutday(i,2)=vday(i+30-vdays,2)
		else
			voutday(i,0)=0
			voutday(i,1)=""
			voutday(i,2)=0
		end if
	next
end if
%>
            <td align=right valign=top><p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 13"> <font face="Arial"><%=int(maxday*10+0.5)/10%></font>
                <p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 13"> <font face="Arial"><%=int(3*maxday*10/4+0.5)/10%></font>
                <p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 13"> <font face="Arial"><%=int(maxday*10/2+0.5)/10%></font>
                <p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 0"> <font face="Arial"><%=int(maxday*10/4+0.5)/10%><br>
              </font></td>
            <td width=10><img src="images/tu_back_left.gif"></td>
            <%
'
for i= 0 to 30
%>
            <td width=13 valign=bottom background="images/tu_back.gif" align=center><img style="BORDER-BOTTOM: #000000 1px solid" src="images/tu.gif"
	height="<%=(voutday(i,0)/maxday)*100%>" width="9"
	alt="<%=voutday(i,1)%>�գ�����<%=voutday(i,0)%>�Σ�<%
	'����������İٷ�������ȷ��С����1λ��С�������ǰ�����ĸ0
	lsbf=int(voutday(i,0)*1000/sumday+0.5)/10
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
for i= 0 to 30
%>
            <td width=13 align=center><a title="<%=voutday(i,1)%>�գ�����<%=voutday(i,0)%>�Σ�<%
	'����������İٷ�������ȷ��С����1λ��С�������ǰ�����ĸ0
	lsbf=int(voutday(i,0)*1000/sumday+0.5)/10
	if lsbf<1 then lsbf="0" & lsbf
	Response.Write lsbf
	%>%">
              <%
'���ݵ���������ò�ͬ����ɫ�������ڣ���������ɫ�������ú�ɫ
select case voutday(i,2)
case 1%>
              <font face="Arial" style="letter-spacing: -1" color="red">
              <%case 7%>
              <font face="Arial" style="letter-spacing: -1" class="fonts">
              <%case else%>
              <font face="Arial" style="letter-spacing: -1">
              <%end select%>
              <%=voutday(i,1)%></font></font></font></a></td>
            <%
next
%>
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
    <td width="498">�����·��շ�����<br>
        <table border="0" cellpadding="0" cellspacing="0" width="453">
          <tr height="9">
            <td></td>
          </tr>
          <tr height="101">
            <%
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select vday,count(id) as allday from view group by vday"
rs.Open sql,conn,1,1

dim vallday(31)
maxallday=0
sumallday=0
do while not rs.EOF
	vallday(clng(rs("vday"))-1)=clng(rs("allday"))
	if vallday(clng(rs("vday"))-1)>maxallday then maxallday=vallday(clng(rs("vday"))-1)
	sumallday=sumallday+vallday(clng(rs("vday"))-1)
	rs.MoveNext
loop
'��ֹ����Ϊ0������
if maxallday=0 then maxallday=1
if sumallday=0 then sumallday=1
%>
            <td align=right valign=top><p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 13"> <font face="Arial"><%=int(maxallday*10+0.5)/10%></font>
                <p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 13"> <font face="Arial"><%=int(3*maxallday*10/4+0.5)/10%></font>
                <p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 13"> <font face="Arial"><%=int(maxallday*10/2+0.5)/10%></font>
                <p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 0"> <font face="Arial"><%=int(maxallday*10/4+0.5)/10%><br>
              </font></td>
            <td width=10><img src="images/tu_back_left.gif"></td>
            <%
for i= 0 to 30
%>
            <td width=13 valign=bottom background="images/tu_back.gif" align=center><img style="BORDER-BOTTOM: #000000 1px solid;" src="images/tu.gif"
	height="<%=(vallday(i)/maxallday)*100%>" width="9"
	alt="<%=i+1%>�գ�����<%=vallday(i)%>�Σ�<%
	'����������İٷ�������ȷ��С����1λ��С�������ǰ�����ĸ0
	lsbf=int(vallday(i)*1000/sumallday)/10
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
for i= 0 to 30
%>
            <td width=13 align=center><a title="<%=i+1%>�գ�����<%=vallday(i)%>�Σ�<%
	'����������İٷ�������ȷ��С����1λ��С�������ǰ�����ĸ0
	lsbf=int(vallday(i)*1000/sumallday)/10
	if lsbf<1 then lsbf="0" & lsbf
	Response.Write lsbf
	%>%">
              <%
	'��Ϊ���ڵ����ڿ���̫���������ø�һ��ʾ�İ취
	if (i+1) mod 2 then%>
              <font face="Arial" style="letter-spacing: -1"><%=i+1%></font>
              <%end if%>
            </a></td>
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
'����ָ�����ڵķ�����
function vdaycon(theday)
    theday=cdate(theday)
    thetday=cdate(theday+1)
    tmprs=conn.execute("Select count(id) as vdaycon from view where" & _
		" vtime>=datevalue('" & theday & "') and vtime<=datevalue('" & thetday & "')")
    vdaycon=tmprs("vdaycon")
    set tmprs=nothing
	if isnull(vdaycon) then vdaycon=0
end function
%>