<!-- #include file="iAccessRight.asp" -->
<!-- #include file="chkAdminPas.asp" -->
<!-- #include file="conn.asp" -->

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
<p align="center"><span style="font-size: 12pt">用户管理·修改用户类别</span></p>

<div align="center">
        
  <table width="90%" border="1" cellpadding="0" cellspacing="0" bordercolordark="#FFFFFF">
    <tr bgcolor="#CCCCCC"> 
      <td width="120" height="25" align="center">类别名</td>
      <td width="80" align="center">用户权限ID</td>
      <td align="center">类别简介</td>
      <td width="80" align="center">修改</td>
    </tr>
    <%
set rs=conn.execute("select * from userTypeInfo")
%>
    <%do while not rs.eof%>
    <tr> 
      <td height="22" align="center"><%=rs("typeName")%></td>
      <td align="center"><%=rs("typeId")%></td>
      <td> <%=rs("typeInfo")%> </td>
      <td align="center"><a href="userTypeModify.asp?typeid=<%=rs("typeid")%>">修改</a></td>
    </tr>
    <%
rs.movenext
loop
%>
  </table>
</div>
</body>
</html>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
%>

