<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!-- #include file="inc_config.asp" -->
<%
'权限分配
if session.Contents("master")=false and whatcan<1 then Response.Redirect "help.asp?id=004&error=本统计系统管理员不允许访客查看任何信息。"

'所有总访问数、开始访问日期（从主数据库读取）
    tmprs=conn.execute("Select vtop,starttime from vjian")
    vtotal=tmprs("vtop")
    vfirst=tmprs("starttime")
	set tmprs=nothing
	if isnull(vtotal) then vtotal=0

if vtotal=0 then
	conn.Close
	set conn=nothing
	Response.Redirect "help.asp?id=007&error=统计系统还没有启用，尚不能查看统计报告。"
end if


'在线人数
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql="select vip from view where vtime >= dateadd('n',-20,now()) group by vip"
	rs.Open sql,conn,1,1
	vonline=rs.RecordCount
	rs.Close
	set rs=nothing
'今日访问量、昨日访问量（从简数据朝读取）
    tmprs=conn.execute("Select today,yesterday from vjian")
    vtoday=tmprs("today")
    vyesterday=tmprs("yesterday")
	if isnull(vtoday) then vtoday=0
	if isnull(vyesterday) then vyesterday=0
'今年访问量
    tmprs=conn.execute("Select count(id) as vthisyear from view where vyear=" & year(now))
    vthisyear=tmprs("vthisyear")
	if isnull(vthisyear) then vthisyear=0
'本月访问量
    tmprs=conn.execute("Select count(id) as vthismonth from view where vmonth=" & month(now) & " and vyear=" & year(now))
    vthismonth=tmprs("vthismonth")
	if isnull(vthismonth) then vthismonth=0
'访问天数、平均每天访问量
	vdays=now()-vfirst
	vdayavg=vtotal/vdays
	vdays=int((vdays*10^mPrecision)+0.5)/(10^mPrecision)
	if vdays<1 then vdays="0" & vdays
	vdayavg=int((vdayavg*10^mPrecision)+0.5)/(10^mPrecision)
'预计今日访问量
	vdaylong=now()-date()
	vguess=int(((vtoday/vdaylong)+vyesterday)/2+0.5)
	if vguess< vtoday then vguess=int((vtoday/vdaylong)+0.5)
'当前用户放量
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
    <td bgcolor="#CCCCCC">总体数据 </td>
  </tr>
  <tr>
    <td>总访问量: &nbsp; <%=vtotal+old_count%><br>
用本系统前: <%=old_count%><br>
用本系统后: <%=vtotal%><br>
在线人数: &nbsp; <%=vonline%><br>
您的访问量: <%=vuser%><br>
开始统计于: <%=vfirst%><br>
今日访问量: <%=vtoday%><br>
昨日访问量: <%=vyesterday%><br>
今年访问量: <%=vthisyear%><br>
本月访问量: <%=vthismonth%><br>
统计天数: &nbsp; <%=vdays%><br>
平均日访量: <%=vdayavg%><br>
预计今日: &nbsp; <%=vguess%></td>
  </tr>
</table>
</body>
</html>
