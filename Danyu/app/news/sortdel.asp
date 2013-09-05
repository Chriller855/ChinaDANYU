<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!-- #include file="conn.asp" -->
<%
sortid=clng(request("sortid"))
if sortid=1 then
	response.write "<script language=javascript>alert('不能删除总类！');location.href='editlanmu.asp'</script>"
else
   dim rs,sql
  
   set rs=server.createobject("adodb.recordset")
   sql="delete from newssort where sortid="&sortid&" or attributeid="&sortid
   response.write sql
   rs.open sql,conn,3,4 
   conn.close
   set conn=nothing
   response.redirect "editlanmu.asp"
end if
%>