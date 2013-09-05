<!-- www.00ok.net 版权所有 -->
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
<td colspan="6" align="right">登记时间：<%=rs("dateAndTime")%></td></tr>
</table>
<table id="jlb" width="650" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#e5e5e5" style="border: 4 solid #e5e5e5;border-collapse:collapse;">
  <tr>
    <td><strong>姓名</strong></td>
    <td><%=rs("memberName")%></td>
    <td><strong>性别</strong></td>
    <td><%=rs("postCode")%></td>
    <td><strong>年龄</strong></td>
    <td><%=rs("address")%></td>
  </tr>
  <tr>
    <td><strong>学历</strong></td>
    <td><%=rs("ok")%></td>
    <td><strong>应聘职位</strong></td>
    <td><%=rs("comeDate")%></td>
    <td><strong>工作经验</strong></td>
    <td><%=rs("ok2")%></td>
  </tr>
  <tr>
    <td><strong>电话</strong></td>
    <td><%=rs("telphone")%></td>
    <td><strong>手机</strong></td>
    <td><%=rs("mobilePhone")%></td>
    <td><strong>电子邮件</strong></td>
    <td><%="<a href='mailto:"&rs("email")&"'>"&rs("email")&"</a>"%></td>
  </tr>
  <tr>
    <td colspan="6"><strong>个人工作简历</strong></td>
  </tr>
  <tr>
    <td colspan="6"><%=rs("comment")%></td>
  </tr>
  <tr>
    <td colspan="6" align="right"><A href="javascript:window.print();">[打印]</A>[<a href="javascript:history.back();">返回</a>]</td>
  </tr>
</table>
<p>&nbsp;</p>
</body>
</html>
<%
conn.close
set conn=nothing
%>