<!-- www.00ok.net 版权所有 -->
<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!--#include file="conn.asp"-->
<%
if Request("act") = "make" then

dim rs,sql,filename,fs,myfile,x,link,fileURL

Set fs = server.CreateObject("scripting.filesystemobject")
'--假设你想让生成的EXCEL文件做如下的存放
'filename = "d:\temp\online.xls"
fileURL="../CRlandClubmemberExcel.xls"
filename = Server.MapPath(fileURL)
'--如果原来的EXCEL文件存在的话删除它
if fs.FileExists(filename) then
fs.DeleteFile(filename)
end if
'--创建EXCEL文件
set myfile = fs.CreateTextFile(filename,true)



Set rs = Server.CreateObject("ADODB.Recordset")
'--从数据库中把你想放到EXCEL中的数据查出来
sql = "select memberName,sexid,birthday,Nationality,placeOfBirth,idNum,address,postCode,telHome,Mobile,email,PlaceOfEmployment,occupation,education,IntendedPrice,IntendedRequirements,IntendedUsage,IntendedDistrict,interest,MemberType,BuildingName,BuildingNumber from clubMember order by dateAndTime desc"
rs.Open sql,conn
if rs.EOF and rs.BOF then

else

dim strLine,responsestr
strLine=""
For each x in rs.fields
strLine= strLine & x.name & chr(9)
Next

'--将表的列名先写入EXCEL
myfile.writeline strLine

Do while Not rs.EOF
strLine=""

for each x in rs.Fields
strLine= strLine & x.value & chr(9)
next
'--将表的数据写入EXCEL
myfile.writeline strLine

rs.MoveNext
loop

end if

rs.Close
set rs = nothing
conn.close
set conn = nothing
set myfile = nothing
Set fs=Nothing

'link="<A HREF=" & filename & ">Open The Excel File</a>"
Response.redirect (fileURL)
end if
%>
</BODY>
</HTML>
