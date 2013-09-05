<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!--#include file="conn.asp"-->
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../index.css" rel="stylesheet" type="text/css">
</head>
<body>
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="0">
  <form action="<%=request.ServerVariables("HTTP_REFERER")%>" name=form1>
    <tr> 
      <td width="77" align=right> 姓名：</td>
      <td> <input type=text name=name size="30" value="<%=whichSite%>"></td>
    </tr>
    <tr> 
      <td align="right" valign="top"> 留言：</td>
      <td><textarea cols="60" rows=15 name=comment></textarea></td>
    </tr>
    <tr> 
      <td> 
    <tr> 
      <td><input name="subid" type="hidden" value="<%=request("subid")%>"></td>
      <td> <input name="submit" type=submit value="发 送">
        <a href="javascript:history.go(-1)"><font color="#FF0000">&lt;&lt; back</font></a><br></td>
      <td width="1"> 
    <tr> 
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td> 
  </form>
  <tr> 
</table>
</body>
</html>
