<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!-- #include file="conn.asp" -->
<%
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
conn.execute("update newsSort set sortPos="&cint(request("sortPos"))&" where sortid="&clng(request("sortid")))

set conn=nothing
Response.write "<script language=javascript>alert('�޸ĳɹ���');location.href='"&session("refererString")&"';</script>"
%>