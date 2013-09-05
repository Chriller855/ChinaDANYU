<!-- #include file="iAccessRight.asp" -->
<!-- #include file="chkAdminPas.asp" -->
<%
if not session("master") then
response.write "<br><br><br><br><p align=center style=font-size:14>请以管理员账户登录!<br><br><a href="
response.write request.ServerVariables("SCRIPT_NAME")&">返回</a></p>"
session("loginFlag")=""
session("userlevel")=""
response.end
end if
%>
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
	document.location.href='userDelete.asp?id='+id;
}
</script>


<body topmargin="0" leftmargin="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
<p align="center"><span style="font-size: 12pt">用户管理·修改用户</span></p>

<div align="center">
        
  <table width="90%" border="1" cellpadding="0" cellspacing="0" bordercolordark="#FFFFFF">
    <tr bgcolor="#CCCCCC"> 
      <td width="120" height="25" align="center">用户名</td>
      <td align="center">用户权限</td>
      <td width="60" align="center">修改</td>
      <td width="60" align="center">删除</td>
    </tr>
<%
set rs=conn.execute("select * from userInfo")
%>
<%do while not rs.eof%>
    <tr> 
      <td height="22" align="center"><%=rs("userId")%></td>
      <td> 
        <%
			set rs2=conn.execute("select * from userTypeInfo where typeID="&rs("userType")&"")
			if not rs2.eof then
				response.write rs2("typeName")&" - <font color=333388>"&rs2("typeInfo")&"</font>"
				rs2.close
			end if
		  %>
      </td>
      <td align="center"><a href="userModify.asp?id=<%=rs("id")%>">修改</a></td>
      <td align="center"><a href="javascript:cdel('<%=rs("id")%>')">删除</a></td>
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

