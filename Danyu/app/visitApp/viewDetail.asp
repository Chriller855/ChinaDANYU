<!-- www.00ok.net ��Ȩ���� -->
<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!--#include file="conn.asp"-->
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="/index.css" rel="stylesheet" type="text/css">
<style type="text/css">
#jlb td {
	padding-left:5px;
}
</style>
</head>
<body>
<%
set rs = createObject("adodb.recordset")
sql="select * from visitApp where id="&request("id")
rs.open sql,conn,1,1
%>
<table width="650" border="0" align="center" cellpadding="0" cellspacing="0">
<tr>
<td colspan="6" align="right">�Ǽ�ʱ�䣺<%=rs("dateAndTime")%></td></tr>
</table>
<table id="jlb" width="650" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#e5e5e5" style="border: 4 solid #e5e5e5;border-collapse:collapse;">
  <tr>
    <td><strong>����</strong></td>
    <td><%=rs("memberName")%></td>
    <td><strong>�Ա�</strong></td>
    <td><%=rs("postCode")%></td>
    <td><strong>����</strong></td>
    <td><%=rs("address")%></td>
  </tr>
  <tr>
    <td><strong>ѧ��</strong></td>
    <td><%=rs("ok")%></td>
    <td><strong>ӦƸְλ</strong></td>
    <td><%=rs("comeDate")%></td>
    <td><strong>��������</strong></td>
    <td><%=rs("ok2")%></td>
  </tr>
  <tr>
    <td><strong>�绰</strong></td>
    <td><%=rs("telphone")%></td>
    <td><strong>�ֻ�</strong></td>
    <td><%=rs("mobilePhone")%></td>
    <td><strong>�����ʼ�</strong></td>
    <td><%="<a href='mailto:"&rs("email")&"'>"&rs("email")&"</a>"%></td>
  </tr>
  <tr>
    <td colspan="6"><strong>���˹�������</strong></td>
  </tr>
  <tr>
    <td colspan="6"><%=rs("comment")%></td>
  </tr>
  <tr>
    <td colspan="6" align="right"><A href="javascript:window.print();">[��ӡ]</A>[<a href="javascript:history.back();">����</a>]</td>
  </tr>
</table>
<p>&nbsp;</p>
</body>
</html>
<%
conn.close
set conn=nothing
%>