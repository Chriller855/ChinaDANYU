<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<%
''''''''''''sortid 列表'''''''''''''''''''''''
'iUrl="newsadd.asp"
if iUrl="" or isnull(iUrl) then iUrl=request("sortid")

response.write "<a href='"&iUrl&"?sortid=0'><font color=black><b>"&whichSite&" 总类</b></font></a><br><br>"
if request("id")<>"" then
	Set tRs2=Server.CreateObject("Adodb.Recordset")
	Sql="Select * from article where id="&clng(request("id"))
	tRs2.open sql,conn,1,1
	iSortId=tRs2("sortid")
	tRs2.close
	set tRs2=nothing
end if
if request("sortid")<>"" then iSortId=clng(request("sortid"))
call listSort(1,1)
''''''''''''''''''''''''''''''''''''''''''''
sub listSort(sortId,i)
	for ii=1 to i
		iSpace=iSpace+"　"
	next
	Set tRs=Server.CreateObject("Adodb.Recordset")
	Sql="Select * from newssort where attributeId="&sortId&" order by sortPos desc, createtime desc"
	tRs.open sql,conn,1,1
	do while not tRs.eof
		response.write iSpace&"· "&"<a href='"&iUrl&"?sortid="&tRs("sortid")&"'>"
		if i=1 then response.write"<b>"
		response.write tRs("sortname")
		if i=1 then response.write"</b>"
		response.write "</a><br>"
		call listSort(tRs("sortid"),i+1)
		tRs.MoveNext
	loop
	tRs.close
	set tRs=nothing
end sub
%>
