<%iAccessRight=""%>
<!-- #include file="chkAdminPas.asp" -->
<!-- #include file="conn.asp" -->
<!-- #include file="md5.asp" -->
<%
if session("loginFlag")="" then
	response.write "<br><br><br><br><p align=center style=font-size:14>URL ����!<br><br><a href="
	response.write "javascript:history.go(-1)>����</a></p>"
	response.End
end if

if request("action")="modify" then
	if trim(replace(request("userPwd"),"'",""))="" then
		response.write "<br><br><br><br><p align=center style=font-size:14>����������!<br><br><a href="
		response.write "javascript:history.go(-1);>����</a></p>"
	else
		sql="update userInfo set userPwd='"&md5(trim(replace(request("userPwd"),"'","")))&"' where userId='"&session("loginFlag")&"'"	
		conn.execute(sql)	
		response.write "<br><br><br><br><p align=center style=font-size:14>�޸ĳɹ�!<br><br><a href="
		response.write "userModifyPWD.asp>����</a></p>"
	end if
	response.End
end if
set rs=conn.execute("select * from userInfo where userId='"&session("loginFlag")&"'")
%>

<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�û������޸��û�</title>
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
<p align="center"><span style="font-size: 12pt">�û������޸�����</span></p>

<form action="<%=request.ServerVariables("SCRIPT_NAME")%>?action=modify" method="POST" name="userForm" onSubmit="return checkPWD();">
    <div align="center">
        
    <script language="JavaScript">
	function checkPWD(){
		if(userForm.userPwd.value!=userForm.userPwdConfirm.value){
			 alert('������������벻һ�£����������룡��');
			 userForm.userPwd.value=userForm.userPwdConfirm.value='';
			 userForm.userPwd.focus();
			 return false;
		}
		return true;
	}</script>
		<table width="388" border="0" cellspacing="0" cellpadding="0">
      <tr> 
        <td width="153" align="right">�û�����</td>
        <td width="235"> <input type="hidden" name="userId" value="<%=rs("userId")%>">
          <span style="font-size: 10pt;font-color:#cc0000"><font color=#cc0000><b><%=rs("userId")%></b></font> </span></td>
      </tr>
      <tr> 
        <td align="right">���룺</td>
        <td><input type="password" name="userPwd" size="20" style="font-size: 9pt; font-family: ���� (serif)"></td>
      </tr>
      <tr>
        <td align="right">����ȷ�ϣ�</td>
        <td><input type="password" name="userPwdConfirm" size="20" style="font-size: 9pt; font-family: ���� (serif)"></td>
      </tr>
    </table>
    
  <p><span style="font-size: 9pt">&nbsp;&nbsp; 
    <input type="submit" value="��  ��"
  name="B1">
      &nbsp; </span></p></div>
</form>
</body>
</html>
<%
conn.close
set conn=nothing
%>

