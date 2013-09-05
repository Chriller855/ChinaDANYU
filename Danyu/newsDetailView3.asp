<!--#include file="app/news/conn.asp" -->
<html>

<body>
<%
sql="select top 1 * from article where id="& clng(request("id"))
'response.write sql
Set Rs=Server.CreateObject("Adodb.Recordset")
rs.open sql,conn,1,1
response.write "<font style='font-size:12px;color=#6b6b6b;'>"
response.write rs("news_content")
response.write "</font>"
rs.close
set rs=nothing
conn.close
set conn=nothing
%>

</body>
</html>