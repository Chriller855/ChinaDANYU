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
<!-- #include file="conn.asp" -->

<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�û������޸��û�</title>
<link rel="stylesheet" type="text/css" href="../index.css">
</head><style type="text/css">
td, p, div, br{FONT-SIZE: 12px;Color:#000000;}
</style>
<script language="JavaScript">
function cdel(id){
if (confirm("���Ҫɾ����"))
	document.location.href='userDelete.asp?id='+id;
}
</script>


<body topmargin="0" leftmargin="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
<p align="center"><span style="font-size: 12pt">�û������޸��û�</span></p>

<div align="center">
        
  <table width="90%" border="1" cellpadding="0" cellspacing="0" bordercolordark="#FFFFFF">
    <tr bgcolor="#CCCCCC"> 
      <td width="120" height="25" align="center">�û���</td>
      <td align="center">�û�Ȩ��</td>
      <td width="60" align="center">�޸�</td>
      <td width="60" align="center">ɾ��</td>
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
      <td align="center"><a href="userModify.asp?id=<%=rs("id")%>">�޸�</a></td>
      <td align="center"><a href="javascript:cdel('<%=rs("id")%>')">ɾ��</a></td>
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

