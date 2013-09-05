<!--#include file="../connstr.asp"-->
<%
	'ON ERROR RESUME NEXT
	'db="/app/clubMember/young_user_db.mdb"   '用户数据库
	'connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(db)
	Set conn = Server.CreateObject("ADODB.Connection")
	conn.Open connstr
%>