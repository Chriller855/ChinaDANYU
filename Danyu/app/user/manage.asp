<!-- #include file="iAccessRight.asp" -->
<!-- #include file="chkAdminPas.asp" -->
<%
if not session("master") then
response.write "<br><br><br><br><p align=center style=font-size:14>���Թ���Ա�˻���¼!<br><br><a href="
response.write request.ServerVariables("SCRIPT_NAME")&">����</a></p>"
session("loginFlag")=""
session("userlevel")=""
response.end
end if
%>
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�û�����</title>
<style type="text/css">
td, p, div, br{FONT-SIZE: 12px;Color:#000000;}
</style>
<link rel="stylesheet" type="text/css" href="../index.css">
</head>

<body topmargin="0" leftmargin="0">
<table width="100%" border="0" cellspacing="0" cellpadding="4">
  <tr> 
    <td bgcolor="#CCCCCC" style="FONT-SIZE: 12pt;Color:#000000;">�û�����</td>
  </tr>
  <tr> 
    <td><a href="userAdd.asp" target="right">�����û�</a></td>
  </tr>
  <tr> 
    <td><a href="userModifyList.asp" target="right">�޸ġ�ɾ���û�</a></a></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td><a href="userTypeAdd.asp" target="right">�����û����</a></td>
  </tr>
  <tr>
    <td><a href="userTypeModifyList.asp" target="right">�޸��û����</a></a></td>
  </tr>
  <tr>
    <td><a href="quit.asp" target="_parent">�˳�</a></td>
  </tr>
</table>
<p align="center">&nbsp;</p>

<p align="center">&nbsp;</p>

</body>
</html>