<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!-- #include file="inc_config.asp" -->
<%
'获取条件
offtime=Request("offtime")
if (not isdate(offtime)) then response.end() 'Response.Redirect "help.asp?error=请正确填写要备份数据的截止日期。"

'创建数据对象
'conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBPath

set bakconn=server.createobject("adodb.connection")
bakDBPath = Server.MapPath(bakconnpath)
bakconn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & bakDBPath
Set bakrs = Server.CreateObject("ADODB.Recordset")

'从主库中提取每日访问量
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select count(id) as tcid,datevalue(vtime) as tdate from view where vtime<datevalue('" & offtime & "') and bakdays=0 group by datevalue(vtime)"
rs.Open sql,conn,1,1

Set tmprs = Server.CreateObject("ADODB.Recordset")

bakok=false		'初始备份完成标记为未完成
for i=1 to 2
	if rs.EOF then
		bakok=true
		exit for
	end if
	'计算当日的IP流量
	sql="Select count(id) as abc from view where datevalue(vtime)=datevalue('" & rs("tdate") & "') group by vip"
	tmprs.Open sql,conn,1,1
	tcip	=tmprs.RecordCount
	tcid	=rs("tcid")
	tdate	=rs("tdate")
	tmprs.Close
	'Response.Write rs("tcid") & "," & tcip & "," & rs("tdate") & "<br>"
	'将当前行追加到后备库
	sql="select * from days where datevalue(tdate)=datevalue('" & tdate & "')"
	bakrs.Open sql,bakconn,3,2
	if bakrs.EOF then	'如果后备库中没有这一天
	bakrs.AddNew 
	bakrs("tdate")=tdate
	bakrs("tview")=tcid
	bakrs("tip")=tcip
	bakrs.Update 
	else		'如果已经有这一天了，就追加数据
	bakrs("tview")=bakrs("tview")+tcid
	bakrs("tip")=bakrs("tip")+tcip
	bakrs.Update 
	end if
	bakrs.Close
	
	'将当前日期的记录标记为已备份
	conn.execute("update view set bakdays=1 where datevalue(vtime)=datevalue('" & tdate & "')")
	rs.MoveNext
next
	set tmprs=nothing


rs.Close
set rs=nothing
conn.Close
set conn=nothing

set bakrs=nothing
bakconn.Close
set bakconn=nothing
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<title><%=countname%>－数据备份－第二步－备份每日数据</title>
<%if bakok=false then%><meta http-equiv="refresh" content="1; url='bak_step2_in.asp?offtime=<%=offtime%>'"><%end if%>
<link rel="stylesheet" type="text/css" href="../../index.css">
</head>
<body topmargin=5 rightmargin=0 leftmargin=0 vlink=#000000 style="BACKGROUND-IMAGE: none" class=backq>
<br>
<%if bakok then%>
<p class="p1">每日数据备份完成。
<p align="right"><a href='bak_step3.asp?offtime=<%=offtime%>' target="_parent">下一步 备份客户端信息 开始</a>&nbsp;
<%else%>
<a href="bak_step2_in.asp?offtime=<%=offtime%>">页面每自动刷新一次转换2天的数据，每次刷新的时间大约是2～30秒，根据数据量大小有所不同。如果超过这个时间还没有自动刷新，请点击这里。</a>
<%end if%>
</body>
</html>
