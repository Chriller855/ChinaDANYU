<!-- #include file="iAccessRight.asp" -->
<!-- #include file="chkAdminPas.asp" -->
<%
if not session("master") then
response.write "<br><br><br><br><p align=center style=font-size:14>请以管理员账户登录!<br><br><a href="
response.write request.ServerVariables("SCRIPT_NAME")&">返回</a></p>"
session("loginFlag")=""
session("userlevel")=""
response.end
end if
%>
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>用户管理</title>
<style type="text/css">
td, p, div, br{FONT-SIZE: 12px;Color:#000000;}
</style>
<link rel="stylesheet" type="text/css" href="../index.css">
</head>

<body topmargin="0" leftmargin="0">
<table width="100%" border="0" cellspacing="0" cellpadding="4">
  <tr> 
    <td bgcolor="#CCCCCC" style="FONT-SIZE: 12pt;Color:#000000;">用户管理</td>
  </tr>
  <tr> 
    <td><a href="userAdd.asp" target="right">增加用户</a></td>
  </tr>
  <tr> 
    <td><a href="userModifyList.asp" target="right">修改、删除用户</a></a></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td><a href="userTypeAdd.asp" target="right">增加用户类别</a></td>
  </tr>
  <tr>
    <td><a href="userTypeModifyList.asp" target="right">修改用户类别</a></a></td>
  </tr>
  <tr>
    <td><a href="quit.asp" target="_parent">退出</a></td>
  </tr>
</table>
<p align="center">&nbsp;</p>

<p align="center">&nbsp;</p>

</body>
</html>