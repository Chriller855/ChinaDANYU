<%
session("master")=false
session("loginFlag")=""
session("userlevel")=""
'response.redirect request.ServerVariables("HTTP_REFERER")
response.write "<script language='javascript'>parent.document.location.href='../default.asp'</script>"
%>