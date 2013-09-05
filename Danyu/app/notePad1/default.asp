<!-- www.00ok.net 版权所有 -->
<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!--#include file="conn.asp"-->
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../index.css" rel="stylesheet" type="text/css">
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
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td align="center">
      <%
if request("submit")="发 送" then
		if request("saveCookies")="on" then
		response.Cookies("YoungCRland").Expires=Date+365
		response.Cookies("YoungCRland")("name")=request("name")
		response.Cookies("YoungCRland")("email")=request("email")
		else
		response.Cookies("YoungCRland")("name")=""
		response.Cookies("YoungCRland")("email")=""
		end if
		sql="select * from book1"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,conn,3,4
		rs.addnew
		if request("subid")="" then
			rs("subid")=0
		else
			rs("subid")=request("subid")
			Page = CLng(Request("Page"))
		end if
		rs("name")=server.HTMLEncode(request("name"))
		rs("content")=server.HTMLEncode(Request("comment"))
		rs("dateandtime")=now()
		rs("ip")=request.ServerVariables("REMOTE_ADDR")
		rs.updatebatch
		rs.close
		'session("notePadUpdateTime")=now()
		'response.Redirect(request.ServerVariables("PATH_INFO"))
		'response.end
end if
'if session("notePadUpdateTime")="" then session("notePadUpdateTime")=cdate("1900-01-01")
session("refererString")=request.ServerVariables("SCRIPT_NAME")&"?page="&Request("Page")
%>
      <%
set rs1=server.createobject("adodb.recordset")
set rs2=server.createobject("adodb.recordset")
'sql="select * from book1 where id="&request("id") & " or subid="&request("id")
sql1="select * from book1 where subid=0 order by dateAndTime desc"
rs1.open sql1,conn,1,1
totalrec=rs1.recordcount
'response.write totalrec
if rs1.eof then
	response.write "<br><br>没有评论！"
	conn.close
	set conn=nothing
	response.end
end if
const maxPerPage=10
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
RS1.MoveFirst
if currentpage > n then currentpage = n
if currentpage<1 then currentpage=1
RS1.Move (currentpage-1) * maxPerPage
session("notePadRefererString")=request.ServerVariables("SCRIPT_NAME")&"?page="&currentpage
%>
      <%

   For iPage = 1 To maxPerPage
    if rs1.eof then exit for

sql2="select * from book1 where id="&rs1("id")&" or subid="&rs1("id")&" order by subid asc ,id desc"
'response.write sql2&"<br>"
rs2.open sql2,conn,1,1
do while not rs2.eof
%>
      <table width=100% border=0 cellpadding=5 cellspacing=0 <%if rs2("subid")<>0 then response.write "bgcolor=#fafafa"%>>
        <%if rs2("subid")=0 then%>
        <tr bgcolor="#999999"> 
          <td height="1"></td>
        </tr>
        <%end if%>
        <tr> 
          <td> <font style="font-size:14px;" color=#111111>&nbsp; </font>
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="300"><font style="font-size:12px;" color=#111111> 
                  <%
		  if rs2("name")="" then
		  	response.write "NoName"
		  else
		    response.write rs2("name")
		  end if%>
                  </font><font color=#226666 style="font-size:10px;"><%="[ "&rs2("dateandtime")&" ]"%></font><font style="font-size:12px;" color=#111111> 
                  <br>
                  &nbsp; 
                  <%
		  if rs2("email")<>"" then
		  response.write "Email：<a href=mailto:"&rs2("email")&" title="&rs2("email")&"><font color=#8c0000 style='font-size:12px;'>"&rs2("email")&"</font></a>"
		  end if
		  if rs2("homepage")<>""then
		  response.write " <a href="&rs2("homepage")&" title="&rs2("homepage")&"><font color=#8c0000  style='font-size:12px;'>[URL]</font></a>"
		  end if
		  %>&nbsp;
<%if rs2("tel")<>"" then response.write "电话:"&rs2("tel")%>&nbsp;
<%if rs2("address")<>"" then response.write "地址:"&rs2("address")%>
                  </font></td>
                <td align="right" valign="top"><font class="small">
                  <%if rs2("subid")=0 then	
		  	response.write " <a href=reply.asp?subid="&rs2("id")&"><font color=#8c0000 class=small>[回复]</font></a>"
			else
			response.write "&nbsp;"
		end if%>
                  <%
if rs2("subid")=0 then
	response.write "&nbsp;<a href=check.asp?id="&rs2("id")&"&check="
	if rs2("validFlag")=0 then
		response.write "1><font color=#0000ff>[未审核]"
	else
		response.write "0><font color=#FF0000>[审核]"
	end if
	response.write "</font></a>&nbsp;"
end if
%>
                  <%
response.write "&nbsp;<a href=javascript:del('"&rs2("id")&"');><font color=#FF0000 class=small>Delete</font></a>&nbsp;"
%>
                  </font></td>
              </tr>
            </table>
            
          </td>
        </tr>
        <tr> 
          <td height="30"><font class="large" ><%=replace(cstr(rs2("content")),vbCrLf,"<br>")%></font></td>
        </tr>
      </table>
      <%
rs2.movenext
loop
rs2.close
rs1.movenext
next
rs1.close
set rs1=nothing
set rs2=nothing
%>
    </td>
  </tr>
  <tr> 
    <td align="right"> 
      <%				  
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
%>
    </td>
  </tr>
</table>
</body>
</html>
