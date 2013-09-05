<!--#include file="../connstr.asp"-->
<%
	'ON ERROR RESUME NEXT
	'db="../crlandBundDB.mdb"   '用户数据库
	'connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(db)
	'connstr="Driver={SQL Server};Server=localhost;database=dongyuan;uid=sa;pwd=1"
	Set conn = Server.CreateObject("ADODB.Connection")
	conn.Open connstr
%>
