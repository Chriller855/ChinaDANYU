<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!-- #include file="inc_config.asp" -->
<%

'****************** 创建数据对象 ******************
Set rs = Server.CreateObject("ADODB.Recordset")

set connip=server.createobject("adodb.connection")
DBPath = Server.MapPath(ipconnpath)
connip.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath
Set rsip = Server.CreateObject("ADODB.Recordset")


if Request("type")="go" then

tall=clng(Request("tall"))
'如果有“启动”标记，则创建up_date字段并全部标记为未更新。
if Request("start")="yes" then
	'on error resume next
	sql="select count(*) as tall from view where vwhere='未知'"
	temprs=conn.Execute(sql)
	tall=temprs("tall")
	set temprs=nothing
	
	sql="ALTER TABLE view ADD COLUMN up_date bit"
	conn.execute(sql)
	
	sql="update view set up_date=0"
	conn.execute(sql)
elseif tall=0 then
	sql="select count(*) as tall from view where vwhere='未知' and up_date=0"
	temprs=conn.Execute(sql)
	tall=temprs("tall")
	set temprs=nothing
end if

'****************** 更新主数据库 ******************
	on error resume next

	sql="select top 50 * from view where vwhere='未知' and up_date=0"
	rs.Open sql,conn,3,2
	if rs.EOF then
	rs.Close
	sql="ALTER TABLE view drop up_date"
	conn.execute(sql)
	Response.Write "完成"
	else
	
	do while not rs.EOF 
	rs("vwhere")=cadd(rs("vip"))
	rs("up_date")=1
	rs.Update
	rs.MoveNext 
	loop
	rs.Close
	
	sql="select count(*) as thavenot from view where vwhere='未知' and up_date=0"
	temprs=conn.Execute(sql)
	thavenot=temprs("thavenot")
	thave=tall-thavenot
	ttlong=int(thave*300/tall)
	
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=countname%>－更新主数据库</title>
<LINK rel="stylesheet" type="text/css" href="../../index.css">
<meta http-equiv="refresh" content="1; url='update.asp?type=go&tall=<%=tall%>'">
</head>
<body>
<center>
    <br>更 新 进 度<br><br>
<table width="304" cellspacing="1" align="center" cellpadding="0" border="0" class=backs>
  <tr>
    <td align=center>
<table width="300" cellspacing="1" align="center" cellpadding="0" border="0" class=backq>
  <tr height=15>
    <td align=center width="<%=ttlong%>" class=backs>
	</td>
    <td align=center width="<%=300-ttlong%>" bgcolor=white>
	</td>
  </tr>
</table>
	</td>
  </tr>
</table>
<br>已更新/所有: <font class=fonts><%=thave%></font>/<%=tall%>
<br>
如果程序没有继续进行，请点击→[<a href=update.asp?type=go&tall=<%=tall%>>继续</a>]
<br>如果要暂停这一次更新，请点击→[<a href="master.asp">暂停</a>]
<br><br>登录进入更新页面，点击“继续上次未完成的更新”即可继续本次更新。
</center>
</body>
</html>
<%	
end if


else	'如果type<>go
%>
本页功能是用新的IP库更新主记录库，请从管理员页面进入，误闯无罪。

<%
end if

'****************** 关闭数据对象 ******************
set rs=nothing
conn.Close 
set conn=nothing

set rsip=nothing
connip.Close
set connip=nothing

'****************** 自定义函数 ********************
function cadd(sip)
	ipnow=cip(sip)

	sql="select * from ip where (onip<=" & ipnow & " and offip>=" & ipnow & ")"
	rsip.Open sql,connip,1,1

	if rsip.EOF then
	  cadd="未知"
	else
	  cadd=rsip("addj")
	end if
	rsip.Close
end function

function cip(sip)
	tip=cstr(sip)
	sip1=left(tip,cint(instr(tip,".")-1))
	tip=mid(tip,cint(instr(tip,".")+1))
	sip2=left(tip,cint(instr(tip,".")-1))
	tip=mid(tip,cint(instr(tip,".")+1))
	sip3=left(tip,cint(instr(tip,".")-1))
	sip4=mid(tip,cint(instr(tip,".")+1))
	cip=cint(sip1)*256*256*256+cint(sip2)*256*256+cint(sip3)*256+cint(sip4)
end function

function islast()
	islast=true
	on error resume next
	sql="select * from view where up_date=0"
	conn.execute(sql)
	if err<>0 then islast=false
end function
%>