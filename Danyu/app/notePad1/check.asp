<!-- www.00ok.net °æÈ¨ËùÓÐ -->
<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!--#include file="conn.asp"-->
<%
sql="update book1 set validFlag="&request("check")&" where id="&request("id")
response.write sql
'response.end
conn.execute(sql)
conn.close
set conn=nothing
response.redirect session("notePadRefererString")
%>
