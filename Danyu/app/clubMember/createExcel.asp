<!-- www.00ok.net ��Ȩ���� -->
<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!--#include file="conn.asp"-->
<%
if Request("act") = "make" then

dim rs,sql,filename,fs,myfile,x,link,fileURL

Set fs = server.CreateObject("scripting.filesystemobject")
'--�������������ɵ�EXCEL�ļ������µĴ��
'filename = "d:\temp\online.xls"
fileURL="../CRlandClubmemberExcel.xls"
filename = Server.MapPath(fileURL)
'--���ԭ����EXCEL�ļ����ڵĻ�ɾ����
if fs.FileExists(filename) then
fs.DeleteFile(filename)
end if
'--����EXCEL�ļ�
set myfile = fs.CreateTextFile(filename,true)



Set rs = Server.CreateObject("ADODB.Recordset")
'--�����ݿ��а�����ŵ�EXCEL�е����ݲ����
sql = "select memberName,sexid,birthday,Nationality,placeOfBirth,idNum,address,postCode,telHome,Mobile,email,PlaceOfEmployment,occupation,education,IntendedPrice,IntendedRequirements,IntendedUsage,IntendedDistrict,interest,MemberType,BuildingName,BuildingNumber from clubMember order by dateAndTime desc"
rs.Open sql,conn
if rs.EOF and rs.BOF then

else

dim strLine,responsestr
strLine=""
For each x in rs.fields
strLine= strLine & x.name & chr(9)
Next

'--�����������д��EXCEL
myfile.writeline strLine

Do while Not rs.EOF
strLine=""

for each x in rs.Fields
strLine= strLine & x.value & chr(9)
next
'--���������д��EXCEL
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
