<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!-- #include file="conn.asp" -->
<%
conn.execute("delete from article where id="&clng(request("id")))
set conn=nothing
Response.Redirect session("refererString")
%>