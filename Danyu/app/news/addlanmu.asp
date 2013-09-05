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
<p align="center"><span style="font-size: 9pt">新闻管理系统</span></p>
    </td>
  </tr>
  <tr>
    <td width="100%">
    <p align="center"><span style="font-size: 9pt"><br>
    添加栏目</span></p>
    <form method="POST" action="lanmuadd.asp" target="_self">
      <p align="center"><span style="font-size: 9pt">添加<!-- #include file="listTitlesSort.asp" -->类下:<input type="text" name="sortname" size="33">栏目 
      <input type="submit" value="加入" name="B1"><input type="reset" value="重置" name="B2"></span></p>
    </form>
    <p>　</td>
  </tr>
  </table>
</body>
</html>
<% 
Conn.Close
Set Conn=Nothing
%>
