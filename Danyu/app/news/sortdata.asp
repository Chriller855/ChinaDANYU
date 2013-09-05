<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!-- #include file="conn.asp" -->
<%
if clng(request.form("sortid"))=1 then
	response.write "<script language=javascript>alert('不能修改总类！');location.href='editlanmu.asp'</script>"
else
	if request.form("sortid")=request.form("attributeid") then
		response.write "<script language=javascript>alert('不能修改属于自身的类！');history.go(-1);</script>"
	else
		set rs=server.createobject("adodb.recordset")
		sql="select * from newssort where sortid ="&clng(request.form("sortid"))
		rs.open sql,conn,1,3
		
		rs("sortname")=request.form("sortname")
		rs("attributeid")=request.form("attributeid")
		if request.form("updateTime")="1" then rs("createtime")=now()
		
		rs.update
		rs.close
		set rs=nothing
		Conn.Close
		Set Conn=Nothing
		response.write "<script language=javascript>location.href='"&session("refererString")&"';</script>"
	end if
end if
%>
