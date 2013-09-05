<!-- #include file="iAccessRight.asp" -->
<!-- #include file="chkAdminPas.asp" -->
<!-- #include file="conn.asp" -->
<!-- #include file="md5.asp" -->
<%
if request("action")="add" then
	set rs=conn.execute("select id from userinfo where userId='"&trim(request.Form("userId"))&"'")
	if not rs.eof then
		response.write "<br><br><br><br><p align=center style=font-size:14>该用户已存在，请重新输入!<br><br><a href="
		response.write "'javascript:history.go(-1);'>返回</a></p>"
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		response.End
	else
		conn.execute("insert into userInfo (userId,userPwd,userType) values('"&trim(replace(request("userId"),"'",""))&"','"&md5(trim(replace(request("userPwd"),"'","")))&"','"&request.Form("userType")&"')")	
		response.write "<br><br><br><br><p align=center style=font-size:14>添加成功!<br><br><a href="
		response.write request.ServerVariables("SCRIPT_NAME")&">返回</a></p>"
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
<p align="center"><span style="font-size: 12pt">用户管理·添加用户</span></p>

<form action="<%=request.ServerVariables("SCRIPT_NAME")%>?action=add" method="POST" name="userForm" onSubmit="return checkForm();">
    <div align="center">
    <script language="JavaScript">
	function checkForm(){
		if(userForm.userId.value==''){ alert('用户名不能为空！！');userForm.userId.focus();return false;}
		if(userForm.userPwd.value==''){ alert('密码不能为空！！');userForm.userPwd.focus();return false;}
		if(userForm.userPwd.value!=userForm.userPwdConfirm.value){
			 alert('两次输入的密码不一致，请重新输入！！');
			 userForm.userPwd.value=userForm.userPwdConfirm.value='';
			 userForm.userPwd.focus();
			 return false;
		 }
		return true;
	}
	</script>
    <table width="388" border="0" cellspacing="0" cellpadding="0">
      <tr> 
        <td width="153" align="right">用户名：</td>
        <td width="235"><span style="font-size: 9pt"> 
          <input type="text" name="userId" size="20" style="font-size: 9pt; font-family: 宋体 (serif)">
          </span></td>
      </tr>
      <tr> 
        <td align="right">密码：</td>
        <td><input type="password" name="userPwd" size="20" style="font-size: 9pt; font-family: 宋体 (serif)"></td>
      </tr>
      <tr> 
        <td align="right">密码确认：</td>
        <td><input type="password" name="userPwdConfirm" size="20" style="font-size: 9pt; font-family: 宋体 (serif)"></td>
      </tr>
      <tr> 
        <td align="right">用户权限：</td>
        <td> <%
			response.write "<select name='userType'>"
			set rs2=conn.execute("select * from userTypeInfo")
			do while not rs2.eof
				response.write "<option value="&rs2("typeID")&">"&rs2("typeName")&" - "&rs2("typeInfo")&"</option>"
				rs2.movenext
			loop
            response.write "</select>"
		  %></td>
      </tr>
    </table>
    
  <p><span style="font-size: 9pt">&nbsp;&nbsp; 
    <input type="submit" value="提  交"
  name="B1">
      &nbsp; </span></p></div>
</form>
</body>
</html>
