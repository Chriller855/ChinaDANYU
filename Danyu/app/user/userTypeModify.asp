<!-- #include file="iAccessRight.asp" -->
<!-- #include file="chkAdminPas.asp" -->
<!-- #include file="conn.asp" -->
<%

if request("action")="modify" then
	if trim(request("typeName"))<>"" then
		sql= "update userTypeInfo set typeName='"&trim(replace(request("typeName"),"'",""))&"', typeInfo='"&trim(replace(request("typeInfo"),"'",""))&"',typeUrl='"&trim(replace(request("typeUrl"),"'",""))&"' where typeID="&request("typeID")
		'response.write sql
		conn.execute(sql)	
		response.write "<br><br><br><br><p align=center style=font-size:14>�޸ĳɹ�!<br><br><a href="
		response.write "usertypemodifylist.asp>����</a></p>"
		response.End
	else
		response.write "<br><br><br><br><p align=center style=font-size:14>�����������!<br><br><a href="
		response.write "'javascript:history.go(-1);'>����</a></p>"
		response.End
	end if
end if
set rs=conn.execute("select * from userTypeInfo where typeID="&request("typeID"))
%>

<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�û���������û�</title>
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
<p align="center"><span style="font-size: 12pt">�û������޸��û����</span></p>

<form action="<%=request.ServerVariables("SCRIPT_NAME")%>?typeID=<%=request("typeid")%>&action=modify" method="POST" name="userForm" id="userForm">
    <div align="center">
        <table width="388" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="105" align="right">�������</td>
            <td width="283"><span style="font-size: 9pt"> 
              <input name="typeName" type="text" id="typeName" style="font-size: 9pt; font-family: ���� (serif)" size="30" value="<%=rs("typeName")%>">
              </span></td>
          </tr>
          <tr> 
            <td align="right">�����ܣ�</td>
            <td><input name="typeInfo" type="text" id="typeInfo" style="font-size: 9pt; font-family: ���� (serif)" size="30" value="<%=rs("typeInfo")%>"></td>
          </tr>
          <tr> 
            <td align="right">�����ַ��</td>
            <td><input name="typeUrl" type="text" id="typeUrl" style="font-size: 9pt; font-family: ���� (serif)" size="30" value="<%=rs("typeUrl")%>"></td>
          </tr>
        </table>
    
  <p><span style="font-size: 9pt">&nbsp;&nbsp; 
    <input type="submit" value="��  ��"
  name="B1">
      &nbsp; </span></p></div>
</form>
<%conn.close%>
</body>
</html>
