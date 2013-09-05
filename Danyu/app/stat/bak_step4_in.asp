<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!-- #include file="inc_config.asp" -->
<%
'权限检查

'延长脚本超时时间
server.ScriptTimeout =240

'获取条件
offtime=Request("offtime")
if (not isdate(offtime)) then Response.end() 'Response.Redirect "help.asp?error=请正确填写要备份数据的截止日期。"

'创建数据对象
Set rs = Server.CreateObject("ADODB.Recordset")

set bakconn=server.createobject("adodb.connection")
bakDBPath = Server.MapPath(bakconnpath)
bakconn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & bakDBPath
Set bakrs = Server.CreateObject("ADODB.Recordset")

'====================
'    备份来路信息
'====================

'从主库中提取来路信息
sql="select count(id) as tcid,vcome from view where vtime<datevalue('" & offtime & "') and bakpage=0 group by vcome"
rs.Open sql,conn,1,1

do while not rs.EOF 
	'将当前行追加到后备库
	sql="select * from stats where vtype='come' and vtitle='" & rs("vcome") & "'"
	bakrs.Open sql,bakconn,3,2
	if bakrs.EOF then	'如果后备库中没有这一项
		on error resume next
		bakrs.AddNew 
		bakrs("vtype")="come"
		bakrs("vtitle")=rs("vcome")
		bakrs("vdata")=rs("tcid")
		bakrs.Update 
	else		'如果已经有这一项了，就追加数据
		bakrs("vdata")=bakrs("vdata")+rs("tcid")
		bakrs.Update 
	end if
	bakrs.Close
	rs.MoveNext
loop
rs.Close


'====================
'  备份被访页面信息
'====================

'从主库中提取被访页面信息
sql="select count(id) as tcid,vpage from view where vtime<datevalue('" & offtime & "') and bakpage=0 group by vpage"
rs.Open sql,conn,1,1

do while not rs.EOF 
	'将当前行追加到后备库
	sql="select * from stats where vtype='page' and vtitle='" & rs("vpage") & "'"
	bakrs.Open sql,bakconn,3,2
	if bakrs.EOF then	'如果后备库中没有这一项
		bakrs.AddNew 
		bakrs("vtype")="page"
		bakrs("vtitle")=rs("vpage")
		bakrs("vdata")=rs("tcid")
		bakrs.Update 
	else		'如果已经有这一项了，就追加数据
		bakrs("vdata")=bakrs("vdata")+rs("tcid")
		bakrs.Update 
	end if
	bakrs.Close
	rs.MoveNext
loop
rs.Close

'在主库中做上标记，指明这些数据的“客户端信息”已经备份
conn.execute("update view set bakpage=1 where vtime<datevalue('" & offtime & "')")

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
<title><%=countname%>－数据备份－第四步－备份页面信息</title>
<link rel="stylesheet" type="text/css" href="../../index.css">
</head>
<body topmargin=5 rightmargin=0 leftmargin=0 vlink=#000000 style="BACKGROUND-IMAGE: none" class=backq>
<br>
<p class="p1">页面信息备份完成。
<p class="p1" align="right"><a href='bak_step5.asp?offtime=<%=offtime%>' target="_parent">下一步 完成</a>&nbsp;
</body>
</html>
