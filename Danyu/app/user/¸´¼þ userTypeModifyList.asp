<!-- #include file="chkpas.asp" -->
<!-- #include file="chkAdminPas.asp" -->
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
	document.location.href='userTypeDelete.asp?id='+id;
}
</script>


<body topmargin="0" leftmargin="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
<p align="center"><span style="font-size: 12pt">�û������޸��û����</span></p>

<div align="center">
        
  <table width="401" border="0" cellpadding="0" cellspacing="0" bordercolor="#CCCCCC">
    <tr align="center" bgcolor="#CCCCCC"> 
      <td width="178" height="25">�����</td>
      <td width="210">�����</td>
      <td width="210">�޸�</td>
      <td width="210">ɾ��</td>
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
      <td width="210"><a href="userTypeModify.asp?typeid=<%=rs("typeid")%>">�޸�</a></td>
      <td width="210"><a href="javascript:cdel('<%=rs("typeid")%>')">ɾ��</a></td>
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

