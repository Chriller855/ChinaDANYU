<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!-- #include file="conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title></title>
<link href="../index.css" rel="stylesheet" type="text/css">
</head>
<%
dim rs
set rs=server.createobject("adodb.recordset")
sql="select * from newssort where (sortid IS NULL)" 
rs.open sql,conn,1,3
rs.addnew
rs("sortname")=request.form("sortname")
rs("createtime")=now()
rs("sortpos")=1
if request.form("sortid")=0 then
rs("attributeid")=1
else
rs("attributeid")=request.form("sortid")
end if
rs.update
%>
<body>
  <table width="50%" border="1" align="center" cellpadding="0" cellspacing="0" style="border-collapse: collapse">
    <tr bgcolor="#000066"> 
      
    <td width="100%" height="20" align="center" bgcolor="#000000"><font color="#FFFFFF">已添加入数据库</font></td>
    </tr>
    <tr>
      <td width="100%" align="center"> 
        <br>
        栏目名：<%response.write request.form("sortname")%>  <br>
        <br>
    是否继续添加？<br>
          <br>
    <a href="addlanmu.asp">是</a>&nbsp;&nbsp; <a href="manage.asp">否</a><br>
    <br>
        </td>
    </tr>
    </table>
</body>
</html>
<%
rs.close
set rs=nothing
Conn.Close
Set Conn=Nothing
%>
