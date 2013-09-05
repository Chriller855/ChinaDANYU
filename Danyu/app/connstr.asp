<%
	whichSite="[ 丹育 ]"
	'SQL 连接语句
	'connstr="Driver={SQL Server};Server=localhost;database=moonRiver;uid=sa;pwd=00ok"

	'Access 连接语句
	db="/app/#danyudate.mdb"   '用户数据库
	connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(db)
%>