<!-- #include file="user/chkAdminPas.asp" -->
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��Ŀ����</title>
<style type="text/css">
td, p, div, br{FONT-SIZE: 12px;Color:#000000;}
</style>
<link rel="stylesheet" type="text/css" href="index.css">
</head>

<body topmargin="0" leftmargin="0">
<table width="100%" border="0" cellspacing="0" cellpadding="4">
  <tr> 
    <td bgcolor="#CCCCCC"><font  style="FONT-SIZE: 12pt;Color:#000000;">�������</font></td>
  </tr>
  <!-- #include file="News/iAccessRight.asp" -->
  <%if isAccess(iAccessRight) then%>
  <tr> 
    <td><a href="News" target="right0">�������ݹ���</a></td>
  </tr>
  <tr> 
    <td height="8">&nbsp;</td>
  </tr>
  <%end if%>
  <!-- #include file="notePad1/iAccessRight.asp" -->
  <%if isAccess(iAccessRight) then%>
  <tr> 
    <td><a href="notePad1/" target="right0">���Թ���</a></td>
  </tr>
  <tr> 
    <td height="8">&nbsp;</td>
  </tr>
  <%end if%>
  <!-- #include file="visitApp/iAccessRight.asp" -->
  <%if isAccess(iAccessRight) then%>
  <tr> 
    <td><a href="visitApp/" target="right0">���߼�������</a></td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
  </tr>
  <%end if%>
  <!-- #include file="user/iAccessRight.asp" -->
  <%if isAccess(iAccessRight) then%>
  <tr> 
    <td><a href="user/manager.asp" target="right0">�û�����</a></td>
  </tr>
  <%end if%>


  <!-- #include file="stat/iAccessRight.asp" -->
  <%if isAccess(iAccessRight) then%>
  <tr> 
    <td>&nbsp;</td>
  </tr>
  <%end if%>
  <tr> 
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td><a href="user/quit.asp">�˳�</a></td>
  </tr>
</table>
</body>
</html>