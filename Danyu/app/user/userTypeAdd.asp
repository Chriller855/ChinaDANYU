<!-- #include file="iAccessRight.asp" -->
<!-- #include file="chkAdminPas.asp" -->
<%
if not session("master") then
response.write "<br><br><br><br><p align=center style=font-size:14>请以管理员账户登录!<br><br><a href="
response.write request.ServerVariables("SCRIPT_NAME")&">返回</a></p>"
session("loginFlag")=""
session("userlevel")=""
'response.end
end if
%>
<!-- #include file="conn.asp" -->
<%
if request("action")="add" then
	if trim(request("typeName"))<>"" then
		conn.execute("insert into userTypeInfo (typeName,typeInfo,typeUrl,addTime) values('"&trim(replace(request("typeName"),"'",""))&"','"&trim(replace(request("typeInfo"),"'",""))&"','"&trim(replace(request("typeUrl"),"'",""))&"','"&now()&"')")	
		response.write "<br><br><br><br><p align=center style=font-size:14>添加成功!<br><br><a href="
		response.write request.ServerVariables("SCRIPT_NAME")&">返回</a></p>"
		response.End
	else
		response.write "<br><br><br><br><p align=center style=font-size:14>请输入类别名!<br><br><a href="
		response.write "'javascript:history.go(-1);'>返回</a></p>"
		response.End
	end if
end if
%>

<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>用户管理·添加用户</title>
<link rel="stylesheet" type="text/css" href="../index.css">
</head><style type="text/css">
td, p, div, br{FONT-SIZE: 12px;Color:#000000;}
</style>



<body topmargin="0" leftmargin="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
<p align="center"><span style="font-size: 12pt">用户管理·添加用户类别</span></p>

<form action="<%=request.ServerVariables("SCRIPT_NAME")%>?action=add" method="POST" name="userForm" id="userForm">
    <div align="center">
        <table width="388" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="105" align="right">类别名：</td>
            <td width="283"><span style="font-size: 9pt"> 
              <input name="typeName" type="text" id="typeName" style="font-size: 9pt; font-family: 宋体 (serif)" size="30">
              </span></td>
          </tr>
          <tr> 
            <td align="right">类别介绍：</td>
            <td><input name="typeInfo" type="text" id="typeInfo" style="font-size: 9pt; font-family: 宋体 (serif)" size="30"></td>
          </tr>
          <tr> 
            <td align="right">分类地址：</td>
            <td><input name="typeUrl" type="text" id="typeUrl" style="font-size: 9pt; font-family: 宋体 (serif)" size="30"></td>
          </tr>
        </table>
    
  <p><span style="font-size: 9pt">&nbsp;&nbsp; 
    <input type="submit" value="提  交"
  name="B1">
      &nbsp; </span></p></div>
</form>
</body>
</html>
