<!-- #include file="conn.asp" -->
<link href="../index.css" rel="stylesheet" type="text/css">
<%
sql="select * from article where id="& clng(request("id"))
Set Rs=Server.CreateObject("Adodb.Recordset")
rs.open sql,conn,1,1
if rs("news_tpic")<>""then
	response.write "<img src="&rs("news_tpic")&">"
end if
response.write "<br>"
response.write rs("news_title")
response.write "<hr>"
response.write rs("news_SubTitle")
response.write "&nbsp;&nbsp;"
response.write rs("news_author")
response.write "&nbsp;&nbsp;"
response.write year(rs("news_uptime"))&"-"&month(rs("news_uptime"))&"-"&day(rs("news_uptime"))
response.write "<br>"
response.write "<br>"
response.write rs("news_original")
response.write "<br>"
response.write "<br>"
response.write rs("news_content")

rs.close
set rs=nothing
conn.close
set conn=nothing
%>
