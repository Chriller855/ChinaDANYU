<!-- #include file=conn.asp -->
<!-- #include file=md5.asp -->
<%
session.Timeout=20
if not session("master") and session("loginFlag")="" then

	userId=trim(replace(request("userId"),"'",""))
	userPwd=md5(trim(replace(request("userPwd"),"'","")))
		
	Set Rs=Server.CreateObject("Adodb.Recordset")
'	Sql="Select * from userInfo where userid='"&request.form("userid")&"' and userpwd='"&request.form("userpwd")&"'"
	Sql="Select * from userInfo where userid='"&userid&"' and userpwd='"&userpwd&"'"
	
	Rs.open Sql,Conn,1,1
	if rs.eof and rs.bof then
	%>
	
<table width="100%" height="60%">
  <tr> 
    <td align="center" valign="bottom" style='font-size:14;'> <form method=POST action="<%=request.ServerVariables("SCRIPT_NAME")&"?"&request.ServerVariables("QUERY_STRING")%>">
        <p align='center' style='font-size:14;'><font color="#990000">请输入用户名、密码！</font></p>
        <p>用户名： 
          <input type="text" name="userid" size="13">
          <br>
          &nbsp; 密码： 
          <input type="password" name="userpwd" size="13">
        </p>
        <p> 
          <input type="submit" value="登录" name="B1">
        </p>
      </form></td>
  </tr>
  <tr>
  </tr>
</table>
	<%
		session("master")=false
		session("userlevel")=""
		session("loginFlag")=""
		response.end
	else
		if rs("userType")=1 then 
			session("master")=true
		else
			session("master")=false
		end if
		session("userlevel")=rs("userType")
		session("loginFlag")=rs("userid")
		response.Redirect request.ServerVariables("SCRIPT_NAME")&"?"&request.ServerVariables("QUERY_STRING")
	end if
	rs.close
	set rs=nothing
end if
conn.CLOSE
SET conn=NOTHING
%>
