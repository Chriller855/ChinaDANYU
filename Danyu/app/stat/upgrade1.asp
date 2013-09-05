<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!-- #include file="inc_config.asp" -->
<%

'****************** 创建数据对象 ******************

'判断开始统计日期
	tmprs2=conn.execute("Select top 1 vtime from view order by id")
    vfirst=tmprs2("vtime")
    set tmprs2=nothing


'开始升级

if isupgrade() then
	on error resume next
	sql="DROP TABLE ip"
	conn.execute(sql)
end if
if isupgrade12()=false then
	on error resume next
	sql="ALTER TABLE view ADD COLUMN vwidth short"
	conn.execute(sql)
	sql="update view set vwidth=0"
	conn.execute(sql)
end if
if isupgrade13()=false then
	on error resume next
	sql="ALTER TABLE view ADD COLUMN bakdays bit"
	conn.execute(sql)
	sql="ALTER TABLE view ADD COLUMN bakstats bit"
	conn.execute(sql)
	sql="ALTER TABLE view ADD COLUMN bakpage bit"
	conn.execute(sql)
end if
if isupgrade13b()=false then
	on error resume next
	sql="ALTER TABLE vjian ADD COLUMN starttime datetime"
	conn.execute(sql)
	sql="update vjian set starttime='" & vfirst & "'"
	conn.execute(sql)
end if

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<title><%=countname%>－管理员控制台</title>
<link rel="stylesheet" type="text/css" href="../../index.css">
</head>
<body topmargin=5 rightmargin=0 leftmargin=0 vlink=#000000>
<!--#include file="inc_top.asp"-->
<br>
<table width="500" cellspacing="0" align="center" cellpadding="0" border="0">
  <tr><td colspan="3"><img src="images/photoup.gif"></td></tr>
  <tr height="30">
    <td width="1" class="backs"><img src="images/touming.gif" width="1" height="1"></td>
    <td width="498" class="backq">
		&nbsp; <img src="images/tb_index1.gif" align=absmiddle> <font style="font-size:16px">&nbsp;</font>从就版本升级－升级数据库结构<br>
	<%if isupgrade() then%>
	<p class="p1">删除主库中的IP表失败！可能是数据库正在被别的程序调用，无法锁定。
	<%else%>
	<p class="p1">主库IP表删除成功！
	<%end if%>
	<%if isupgrade12() then%>
	<p class="p1">主库view表屏宽字段创建成功！
	<%else%>
	<p class="p1">创建主库屏宽字段失败！可能是数据库正在被别的程序调用，无法锁定。
	<%end if%>
	<%if isupgrade13() then%>
	<p class="p1">主库view表备份标记字段创建成功！
	<%else%>
	<p class="p1">创建主库备份标记字段失败！可能是数据库正在被别的程序调用，无法锁定。
	<%end if%>
	<%if isupgrade13b() then%>
	<p class="p1">在主库vjian表中添加“开始统计日期”字段成功！
	<%else%>
	<p class="p1">在主库vjian表中添加“开始统计日期”字段失败！可能是数据库正在被别的程序调用，无法锁定。
	<%end if%>
	</td>
    <td width="1" class="backs"><img src="images/touming.gif" width="1" height="1"></td>
  </tr>
  <tr><td colspan="4"><img src="images/photodown.gif"></td></tr>
</table>
<br>
<!--#include file="inc_bottom.asp"-->
</body>
</html>
<%
'****************** 关闭数据对象 ******************
conn.Close
set conn=nothing

'****************** 自定义函数 ********************
function isUpgrade()
	isUpgrade=true
	on error resume next
	sql="select * from ip"
	conn.execute(sql)
	if err<>0 then isUpgrade=false
end function

function isUpgrade12()
	isUpgrade12=true
	on error resume next
	sql="select vwidth from view"
	conn.execute(sql)
	if err<>0 then isUpgrade12=false
end function

function isUpgrade13()
	isUpgrade13=true
	on error resume next
	sql="select bakdays from view"
	conn.execute(sql)
	if err<>0 then isUpgrade13=false
end function

function isUpgrade13b()
	isUpgrade13b=true
	on error resume next
	sql="select starttime from vjian"
	conn.execute(sql)
	if err<>0 then isUpgrade13b=false
end function

%>