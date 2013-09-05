<%iAccessRight=""%>
<!-- #include file="chkAdminPas.asp" -->
<!-- #include file="conn.asp" -->
<!-- #include file="md5.asp" -->
<%
if session("loginFlag")="" then
	response.write "<br><br><br><br><p align=center style=font-size:14>URL 出错!<br><br><a href="
	response.write "javascript:history.go(-1)>返回</a></p>"
	response.End
end if

if request("action")="modify" then
	if trim(replace(request("userPwd"),"'",""))="" then
		response.write "<br><br><br><br><p align=center style=font-size:14>请填入密码!<br><br><a href="
		response.write "javascript:history.go(-1);>返回</a></p>"
	else
		sql="update userInfo set userPwd='"&md5(trim(replace(request("userPwd"),"'","")))&"' where userId='"&session("loginFlag")&"'"	
		conn.execute(sql)	
		response.write "<br><br><br><br><p align=center style=font-size:14>修改成功!<br><br><a href="
		response.write "userModifyPWD.asp>返回</a></p>"
	end if
	response.End
end if
set rs=conn.execute("select * from userInfo where userId='"&session("loginFlag")&"'")
%>

<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>用户管理·修改用户</title>
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
<p align="center"><span style="font-size: 12pt">用户管理·修改密码</span></p>

<form action="<%=request.ServerVariables("SCRIPT_NAME")%>?action=modify" method="POST" name="userForm" onSubmit="return checkPWD();">
    <div align="center">
        
    <script language="JavaScript">
	function checkPWD(){
		if(userForm.userPwd.value!=userForm.userPwdConfirm.value){
			 alert('两次输入的密码不一致，请重新输入！！');
			 userForm.userPwd.value=userForm.userPwdConfirm.value='';
			 userForm.userPwd.focus();
			 return false;
		}
		return true;
	}</script>
		<table width="388" border="0" cellspacing="0" cellpadding="0">
      <tr> 
        <td width="153" align="right">用户名：</td>
        <td width="235"> <input type="hidden" name="userId" value="<%=rs("userId")%>">
          <span style="font-size: 10pt;font-color:#cc0000"><font color=#cc0000><b><%=rs("userId")%></b></font> </span></td>
      </tr>
      <tr> 
        <td align="right">密码：</td>
        <td><input type="password" name="userPwd" size="20" style="font-size: 9pt; font-family: 宋体 (serif)"></td>
      </tr>
      <tr>
        <td align="right">密码确认：</td>
        <td><input type="password" name="userPwdConfirm" size="20" style="font-size: 9pt; font-family: 宋体 (serif)"></td>
      </tr>
    </table>
    
  <p><span style="font-size: 9pt">&nbsp;&nbsp; 
    <input type="submit" value="提  交"
  name="B1">
      &nbsp; </span></p></div>
</form>
</body>
</html>
<%
conn.close
set conn=nothing
%>

