<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<%
''''''''''''sortid 列表'''''''''''''''''''''''
response.write "<select size=1 name=attributeid>"
response.write "<option value=1>"&websiteName&" 总类</option>"
call listSort(1,1)
''''''''''''''''''''''''''''''''''''''''''''
sub listSort(isortId,i)
for ii=1 to i
iSpace=iSpace+"&nbsp;&nbsp;"
next
Set tRs=Server.CreateObject("Adodb.Recordset")
Sql="Select * from newssort where attributeId="&isortId&" order by sortPos desc,createtime desc"
tRs.open sql,conn,1,1
do while not tRs.eof
response.write "<option "
if tRs("sortid")=attributeId then
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
