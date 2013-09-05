<!-- www.00ok.net 版权所有 -->
<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!--#include file="conn.asp"-->
<title>在线简历列表</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../../index.css" rel="stylesheet" type="text/css">
<head>
<script language=javascript>
function del(id){
	if(confirm('确定要删除吗？')){
		document.location.href="del.asp?subid="+id;
	}
}
</script>
</head>
<body leftmargin="0" topmargin="4">
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td> 在线简历列表：<br> <%
set rs=server.createobject("adodb.recordset")
sql="select * from visitApp where comeDate<>'wantBook' order by id desc"
rs.open sql,conn,1,1
if rs.eof then
	response.write "暂无简历！"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
totalrec=rs.recordcount
'response.write totalrec
const maxPerPage=20
dim n,p,ii
dim currentpage,totalrec,page_count
currentPage=request("page")
if currentpage="" or not isnumeric(currentpage) then
	currentpage=1
else
	currentpage=clng(currentpage)
end if

if totalrec mod maxPerPage=0 then
	n= totalrec \ maxPerPage
else
	n= totalrec \ maxPerPage+1
end if
RS.MoveFirst
if currentpage > n then currentpage = n
if currentpage<1 then currentpage=1
RS.Move (currentpage-1) * maxPerPage
session("visitAppRefererString")=request.ServerVariables("SCRIPT_NAME")&"?page="&Page
%> 
      <table width="100%" border="1" cellpadding="2" cellspacing="0" bordercolor="#e5e5e5" bordercolordark="#e5e5e5">
        <tr bgcolor="#e5e5e5"> 
          <td>&nbsp;<font color="#000000">姓 名</font></td>
          <td>&nbsp;<font color="#000000">性 别</font></td>
          <td><font color="#000000">年龄</font></td>
          <td> <font color="#000000">学历</font></td>
          <td><font color="#000000">工作经验</font></td>
          <td><font color="#000000">应聘职位</font></td>
          <td><font color="#000000">详情</font></td>
          <td><font color="#000000">删除</font></td>
        </tr>
        <%
do while not rs.eof and page_count<maxPerPage
page_count=page_count+1
%>
        <tr> 
          <td>&nbsp<%=rs("memberName")%></td>
          <td>&nbsp;<%=rs("postCode")%></td>
          <td><%=rs("address")%></td>
          <td><%=rs("ok")%></td>
          <td><%=rs("ok2")%></td>
          <td><%=left(rs("comment"),40)%>&nbsp;...</td>
          <td><%="<a href='viewDetail.asp?id="&rs("id")&"'>详情</a>"%></td>
          <td><%="<a href=javascript:del('"&rs("id")&"');>删除</a>"%></td>
        </tr>
        <%rs.movenext
loop%>
    </table></td>
  </tr>
  <tr> 
    <td align="right"> <%				  
''''''''''''''''''''''
if totalrec<>0 then
response.write "共 <font color=#7c0000><b>"&totalrec&"</b></font> 条 "
	if currentpage-1 mod 10=0 then
		p=(currentpage-1) \ 10
	else
		p=(currentpage-1) \ 10
	end if
	dim pagelist,pagelistbit
	if currentPage=1 then
		response.write "<font face=webdings color=#ff0000>9</font>   "
	else
		response.write "<a href='?sortid="&sortid&"&page=1' title=首页><font face=webdings>9</font></a>   "
	end if
	if p*10>0 then response.write "<a href='?sortid="&sortid&"&page="&Cstr(p*10)&"' title=上十页><font face=webdings>7</font></a>   "
	response.write "<b>"
	for ii=p*10+1 to P*10+10
		if ii=currentPage then
			response.write "<font color=#ff0000>"+Cstr(ii)+"</font> "
		else
			response.write "<a href='?sortid="&sortid&"&page="&Cstr(ii)&"'>"+Cstr(ii)+"</a>   "
		end if
		if ii=n then exit for
		'p=p+1
	next
	response.write "</b>"
	if ii<n then response.write "<a href='?sortid="&sortid&"&page="&Cstr(ii)&"' title=下十页><font face=webdings>8</font></a>   "
	if currentPage=n then
		response.write "<font face=webdings color=#ff0000>:</font>   "
	else
		response.write "<a href='?sortid="&sortid&"&page="&Cstr(n)&"' title=尾页><font face=webdings>:</font></a>"
	end if
end if
''''''''''''''''''''''''''''''''''''''''''
%> </td>
  </tr>
  <tr>
    <td><A href="javascript:window.print();">[打印]</A></td>
  </tr>
</table>
</body>
</html>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
