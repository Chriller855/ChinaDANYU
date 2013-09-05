<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!-- #include file="conn.asp" -->
<%
set rs=server.createobject("adodb.recordset")
searchtxt=request("searchtxt")
level=clng(request("level"))
id=clng(request("id"))
referer=session("refererString")

if level=0 then
level=1 
else
level=0
end if

sql="select * from article where id="&id
rs.open sql,conn,1,3

rs("news_level")=level
rs.update

rs.close 
set rs=nothing
conn.close
set conn=nothing
Response.Redirect  referer
%>                    
