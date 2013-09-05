<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!-- #include file="inc_config.asp" -->
<%
if Request("logpsw")=masterpsw then
	session.Contents("master")=true
	Response.Redirect "index.asp?main=master"
else
	Response.Redirect "help.asp?error=µÇÂ¼ÃÜÂëÊäÈë´íÎó¡£"
end if
%>