<!-- www.00ok.net ��Ȩ���� -->
<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!--#include file="conn.asp"-->
<%
sql="delete from visitApp where id="&request("subid")
conn.execute(sql)
conn.close
set conn=nothing
response.redirect session("visitAppRefererString")
%>
