<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!-- #include file="inc_config.asp" -->
<%
'权限检查
if session.Contents("master")=false and whatcan<2 then Response.Redirect "help.asp?id=004&error=您没有查看日访问统计的权限。"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<title><%=countname%>－日访问统计</title>
<link rel="stylesheet" type="text/css" href="../../index.css">
</head>
<body topmargin=5 rightmargin=0 leftmargin=0 vlink=#000000>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr height="30">
    <td width="498">最近31天访问量<br>
        <table border="0" cellpadding="0" cellspacing="0" width="453">
          <tr height="9">
            <td></td>
          </tr>
          <tr height="101">
            <%
'找到开始统计天数，如果天数不足31天，则跳过前面的空间
tmprs=conn.execute("Select top 1 vtime from view order by id")
vfirst=tmprs("vtime")
set tmprs=nothing

if isnull(vfirst) then vfirst=now()
vdays=int(date()-vfirst+1)

'声明二维数组，voutday(*,0)为访问量,voutday(*,1)为日期,voutday(*,2)为星期
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
	'防止除数为0而出错
	if maxday=0 then maxday=1
	if sumday=0 then sumday=1

'根据已统计天数将数值左移
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
	alt="<%=voutday(i,1)%>日，访问<%=voutday(i,0)%>次，<%
	'计算访问量的百分数，精确到小数后1位，小于零的在前面加字母0
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
            <td width=13 align=center><a title="<%=voutday(i,1)%>日，访问<%=voutday(i,0)%>次，<%
	'计算访问量的百分数，精确到小数后1位，小于零的在前面加字母0
	lsbf=int(voutday(i,0)*1000/sumday+0.5)/10
	if lsbf<1 then lsbf="0" & lsbf
	Response.Write lsbf
	%>%">
              <%
'根据当天的日期用不同的颜色区分日期，周六用绿色，周日用红色
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
    <td width="498">所有月份日访问量<br>
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
'防止除数为0而出错
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
	alt="<%=i+1%>日，访问<%=vallday(i)%>次，<%
	'计算访问量的百分数，精确到小数后1位，小于零的在前面加字母0
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
            <td width=13 align=center><a title="<%=i+1%>日，访问<%=vallday(i)%>次，<%
	'计算访问量的百分数，精确到小数后1位，小于零的在前面加字母0
	lsbf=int(vallday(i)*1000/sumallday)/10
	if lsbf<1 then lsbf="0" & lsbf
	Response.Write lsbf
	%>%">
              <%
	'因为相邻的日期靠的太近，所以用隔一显示的办法
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
用鼠标点指图形柱或者图形柱下的数字可以看到对应的访问量。

<%
rs.Close
set rs=nothing
conn.Close 
set conn=nothing
%>
</body>
</html>
<%
'计算指定日期的访问量
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