<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!-- #include file="conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title></title>
<link href="../index.css" rel="stylesheet" type="text/css">
</head>
<body>

<table border="0" cellspacing="1" width="100%" id="AutoNumber1">
  <tr>
    <td width="100%">
<p align="center"><span style="font-size: 9pt">���Ź���ϵͳ</span></p>
    </td>
  </tr>
  <tr>
    <td width="100%">
    <p align="center"><span style="font-size: 9pt"><br>
    �޸ġ�ɾ����Ŀ</span></p>
    <form method="POST" action="editdata.asp" target="_self">
      <p align="center">
<!-- #include file="listTitlesSort.asp" -->
<%
Conn.Close
Set Conn=Nothing
%>      
      <input type="submit" value="�޸ġ�ɾ��" name="B1"></p>
    </form>
    <p>��</td>
  </tr>
  </table>
</body>

</html>