<!-- www.00ok.net °æÈ¨ËùÓÐ -->
<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!--#include file="conn.asp"-->
<%
'response.write session("notePadRefererString")
sql="delete from clubMember where id="&request("subid")
conn.execute(sql)
conn.close
set conn=nothing
response.redirect session("notePadRefererString")
%>
