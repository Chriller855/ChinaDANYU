<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<%
''''''''''''sortid 列表'''''''''''''''''''''''
response.write "<select size=1 name=sortid>"
response.write "<option value=0>"&whichSite&" 总类</option>"
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
iSpace=iSpace+"&nbsp;&nbsp;"
next
Set tRs=Server.CreateObject("Adodb.Recordset")
Sql="Select * from newssort where attributeId="&sortId&" order by sortPos desc, createtime desc"
tRs.open sql,conn,1,1
do while not tRs.eof
response.write "<option "
'if tRs("sortid")=clng(request("sortid")) then
if tRs("sortid")=iSortId then
response.write "selected "
end if
response.write "value="&tRs("sortid")&">"&iSpace&tRs("sortname")&"</option>"
call listSort(tRs("sortid"),i+1)
tRs.MoveNext
loop
tRs.close
set tRs=nothing
end sub
'''''''''''''''''''''''''''''''''''''''''
response.write "</select>栏目"
'''''''''''''''''''''''''''''''''''''''''
%>
