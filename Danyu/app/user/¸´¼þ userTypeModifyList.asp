<!-- #include file="chkpas.asp" -->
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
<script language="JavaScript">
function cdel(id){
if (confirm("真的要删除吗？"))
	document.location.href='userTypeDelete.asp?id='+id;
}
</script>


<body topmargin="0" leftmargin="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
<p align="center"><span style="font-size: 12pt">用户管理·修改用户类别</span></p>

<div align="center">
        
  <table width="401" border="0" cellpadding="0" cellspacing="0" bordercolor="#CCCCCC">
    <tr align="center" bgcolor="#CCCCCC"> 
      <td width="178" height="25">类别名</td>
      <td width="210">类别简介</td>
      <td width="210">修改</td>
      <td width="210">删除</td>
    </tr>
<%
set rs=conn.execute("select * from userTypeInfo")
%>
<%do while not rs.eof%>
    <tr align="center"> 
      <td height="22"><%=rs("typeName")%></td>
      <td> 
        <%=rs("typeInfo")%>
      </td>
      <td width="210"><a href="userTypeModify.asp?typeid=<%=rs("typeid")%>">修改</a></td>
      <td width="210"><a href="javascript:cdel('<%=rs("typeid")%>')">删除</a></td>
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

