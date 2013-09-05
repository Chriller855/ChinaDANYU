<!-- #include file="iAccessRight.asp" -->
<!-- #include file="chkAdminPas.asp" -->
<!-- #include file="conn.asp" -->
<%
if trim(request("id"))<>"1" then
	conn.execute("delete from userInfo where id="&request("id"))
	conn.close
	set conn=nothing
else
	response.write "<br><br><br><br><p align=center style=font-size:14>超级用户不能被删除!<br><br><a href="
	response.write "'javascript:history.go(-1);'>返回</a></p>"
	response.End
end if
%>
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="refresh" content="1;url=userModifyList.asp">
<title>用户管理·修改用户</title>
<style type="text/css">
td, p, div, br{FONT-SIZE: 12pt;Color:#000000;}
</style>
<link rel="stylesheet" type="text/css" href="../index.css">
</head>
<body topmargin="0" leftmargin="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
<p align="center">删除成功！<br><br>
<a href="userModifyList.asp">返回</a> </p>
</body>
</html>
