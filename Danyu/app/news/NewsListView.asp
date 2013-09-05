<!-- #include file="conn.asp" -->
<%
sortid=100
sql="select * from article where sortid="& sortid &" order by news_uptime desc"
Set Rs=Server.CreateObject("Adodb.Recordset")
rs.open sql,conn,1,1
totalrec=rs.recordcount

dim n,p,ii
dim currentpage,totalrec,page_count
const maxPerPage=20

currentPage=request("page")
if currentpage="" or not isnumeric(currentpage) then
	currentpage=1
else
	currentpage=clng(currentpage)
end if

if totalrec<>0 then
	if totalrec mod maxPerPage=0 then
		n= totalrec \ maxPerPage
	else
	    n= totalrec \ maxPerPage+1
  	end if
	RS.MoveFirst
	if currentpage > n then currentpage = n
    if currentpage<1 then currentpage=1
	RS.Move (currentpage-1) * maxPerPage
dim nn
nn=1


do while not rs.eof and page_count<maxPerPage
page_count=page_count+1

'response.write totalrec-(currentpage-1)*maxPerPage-nn+1 &"  "
if rs("news_Url")="" or isnull(rs("news_Url"))then
	response.write "<a href=newsDetailView.asp?id="&rs("id")&" target=_blank>"
else
	response.write "<a href="&rs("news_Url")&" target=_blank>"
end if
response.write rs("news_title")
response.write "</a>"
response.write "    "
response.write rs("news_uptime")
response.write "<br>"
rs.movenext
nn=nn+1
loop
end if

if totalrec<>0 then
response.write "<br>"
response.write "共有<font color=#ff0000><b>"&totalrec&"</b></font>条 "
if currentpage-1 mod 10=0 then
	p=(currentpage-1) \ 10
else
	p=(currentpage-1) \ 10
end if
dim pagelist,pagelistbit
	if currentPage=1 then
	response.write "<font face=webdings color=#ff0000>9</font>   "
	else
	response.write "<a href='?boardid="&boardid&"&page=1' title=首页><font face=webdings>9</font></a>   "
	end if
	if p*10>0 then response.write "<a href='?boardid="&boardid&"&page="&Cstr(p*10)&"' title=上十页><font face=webdings>7</font></a>   "
	response.write "<b>"
	for ii=p*10+1 to P*10+10
		   if ii=currentPage then
	          response.write "<font color=#ff0000>"+Cstr(ii)+"</font> "
		   else
		      response.write "<a href='?boardid="&boardid&"&page="&Cstr(ii)&"'>"+Cstr(ii)+"</a>   "
		   end if
		if ii=n then exit for
		'p=p+1
	next
	response.write "</b>"
	if ii<n then response.write "<a href='?boardid="&boardid&"&page="&Cstr(ii)&"' title=下十页><font face=webdings>8</font></a>   "
	if currentPage=n then
	response.write "<font face=webdings color=#ff0000>:</font>   "
	else
	response.write "<a href='?boardid="&boardid&"&page="&Cstr(n)&"' title=尾页><font face=webdings>:</font></a>"
	end if
end if

rs.close
set rs=nothing
conn.close
set conn=nothing
%>
