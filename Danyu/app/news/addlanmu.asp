<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!-- #include file="conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title></title>
<link href="../index.css" rel="stylesheet" type="text/css">
</head>
<% 
dim rs
Set Rs=Server.CreateObject("Adodb.Recordset")
Sql="Select * from newssort"
Rs.open Sql,Conn,1,1
%>
<body background="images/bg01.gif">

<table border="0" cellspacing="1" width="100%" id="AutoNumber1">
  <tr>
    <td width="100%">
<p align="center"><span style="font-size: 9pt">���Ź���ϵͳ</span></p>
    </td>
  </tr>
  <tr>
    <td width="100%">
    <p align="center"><span style="font-size: 9pt"><br>
    �����Ŀ</span></p>
    <form method="POST" action="lanmuadd.asp" target="_self">
      <p align="center"><span style="font-size: 9pt">���<!-- #include file="listTitlesSort.asp" -->����:<input type="text" name="sortname" size="33">��Ŀ 
      <input type="submit" value="����" name="B1"><input type="reset" value="����" name="B2"></span></p>
    </form>
    <p>��</td>
  </tr>
  </table>
</body>
</html>
<% 
Conn.Close
Set Conn=Nothing
%>
