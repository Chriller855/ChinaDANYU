<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>公司新闻</title>
<!--<link rel="stylesheet" type="text/css" href="css/base.css">-->
<link rel="stylesheet" type="text/css" href="css/atsunan.css">
<link rel="stylesheet" type="text/css" href="css/news.css">
<link rel="stylesheet" type="text/css" href="index.css">
<script type="text/javascript" src="Scripts/prototype.js" ></script>
<script type="text/javascript" src="Scripts/news_current.js"></script>
<script src="Scripts/AC_RunActiveContent.js" type="text/javascript"></script>
</head>

<body topmargin="0" leftmargin="0" bgcolor="#F5F5F5">

<div align="center">
	<table border="0" cellpadding="0" cellspacing="0" id="table1" width="940" height="80" bgcolor="#FFFFFF">
		<tr>
			<td>
			  <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="940" height="80">
          <param name="movie" value="flash/menu.swf?id=5">
		  <param name="wmode" value="transparent">
          <param name="quality" value="high">
          <embed src="flash/menu.swf?id=5" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="940" height="80">
          </embed>
          </object> 
			</td>
		</tr>
	</table>
</div>
<div align="center">
	<table border="0" cellpadding="0" cellspacing="0" width="940" id="table2" bgcolor="#FFFFFF">
		<tr>
			<td><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=9,0,28,0" width="940" height="184">
			  <param name="movie" value="flash/xtl01.swf">
			  <param name="quality" value="high">
			  <param name="wmode" value="opaque">
			  <embed src="flash/xtl01.swf" quality="high" wmode="opaque" pluginspage="http://www.adobe.com/shockwave/download/download.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="940" height="184"></embed>
		    </object></td>
		</tr>
	</table>
</div>
<div align="center">
	<table border="0" cellpadding="0" cellspacing="0" width="940" id="table3" bgcolor="#FFFFFF" style="border-bottom: 3px solid #E5E5E5">
		<tr>
			<td height="40" colspan="4">　</td>
		</tr>
		<tr>
			<td width="24" valign="top">　</td>
			<td width="197" valign="top">
			<div id="mainLeft">
    <!--左侧菜单开始-->
    <div id="menuOut">
      <div id="menuPor">
<div id="menu"><a href="news01.asp" id="m1"></a> <a href="news02.asp" id="m2"></a><a href="news03.asp" id="m3"></a>
</div>
</div>
    </div>
			</td>
			<td width="28" valign="top">　</td>
			<td width="691" valign="top">
			<div align="center">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table5">
					<tr>
						<td colspan="2">
						<img border="0" src="images/tit0501.gif" width="691" height="47"></td>
					</tr>
					<tr>
						<td width="670">
                        <!--#include file="app/news/conn.asp" -->
                        <% sortid=85 %>
                        <%
if sortid="" or isnull(sortid) then
sortid=request("sortid")
end if
sql="select * from article where sortid="&sortid&" and news_level<>0 order by news_istop desc,news_uptime desc"
Set Rs=Server.CreateObject("Adodb.Recordset")
rs.open sql,conn,1,1
if not rs.eof then
	totalrec=rs.recordcount
	'response.write totalrec
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
	RS.MoveFirst
if currentpage > n then currentpage = n
if currentpage<1 then currentpage=1
RS.Move (currentpage-1) * maxPerPage
%>
       <%
	do while not rs.eof and page_count<maxPerPage
	page_count=page_count+1
%>
                        <div id="news">
                        <div class="pic"><img src="images/dot.gif" width="3" height="5"></div><div class="tit">
						<%
'if rs("sortid")<>2 then										
	if rs("news_URL")="" or isnull(rs("news_URL")) then
	response.write "<a href='newsDetailView.asp?id="&rs("id")&"' target=_parent>"
else
		response.write "<a href='"&rs("news_URL")&"' target=_self>"
	end if
	if rs("news_istop")<>0 then
		response.write "<font color=#c91213>"&rs("news_title")&"</font>"
	else
		tmp=rs("news_title")
		if len(tmp)>35 then
			tmp=left(tmp,60)&"..."
		end if
		response.write "<font color=#6a6a6a>"&tmp&"</font>"
	end if
	response.write "</a>"
'else
'	response.write rs("news_title")
'end if	
%>
<%
if rs("news_tpic")<>"" then response.write "<font color=#c91213>&nbsp;(图)</font>"
%>
</div><div class="dat">
<%response.write year(rs("news_uptime"))&"."
										if len(month(rs("news_uptime")))=1 then response.write "0"
										response.write month(rs("news_uptime"))&"."
										if len(day(rs("news_uptime")))=1 then response.write "0"
										response.write day(rs("news_uptime"))%>
</div>


</div>
<hr class="line">
 <%
rs.movenext
loop
%>
                        
<%
else
	response.write "本栏目暂时没有信息 | Not yet available"
end if%>

<%rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<div class="npage">
<%				  
if totalrec<>0 then
'response.write "共 <font color=#626262><b>"&totalrec&"</b></font> 条 "
	if currentpage-1 mod 10=0 then
		p=(currentpage-1) \ 10
	else
		p=(currentpage-1) \ 10
	end if
	dim pagelist,pagelistbit
	if currentPage=1 then
		response.write "<font face=webdings color=#626262>9</font>   "
	else
		response.write "<a href='?sortid="&sortid&"&page=1' title=首页 target=_self><font face=webdings>9</font></a>   "
	end if
	if p*10>0 then response.write "<a href='?sortid="&sortid&"&page="&Cstr(p*10)&"' title=上十页 target=_self><font face=webdings>7</font></a>   "
	response.write "<b>"
	for ii=p*10+1 to P*10+10
		if ii=currentPage then
			response.write "<font color=#626262>"+Cstr(ii)+"</font> "
		else
			response.write "<a href='?sortid="&sortid&"&page="&Cstr(ii)&"' target=_self><font color=#444444>"+Cstr(ii)+"</font></a>   "
		end if
		if ii=n then exit for
		'p=p+1
	next
	response.write "</b>"
	if ii<n then response.write "<a href='?sortid="&sortid&"&page="&Cstr(ii)&"' title=下十页 target=_self><font face=webdings>8</font></a>   "
	if currentPage=n then
		response.write "<font face=webdings color=#626262>:</font>   "
	else
		response.write "<a href='?sortid="&sortid&"&page="&Cstr(n)&"' title=尾页 target=_self><font face=webdings>:</font></a>"
	end if
end if
%>
</div>
                        </td>
					</tr>
				</table>
			</div>
			</td>
		</tr>
		<tr>
			<td colspan="4">　</td>
		</tr>
	</table>
</div>
<div align="center">
	<table border="0" cellpadding="0" cellspacing="0" width="940" id="table4" bgcolor="#FFFFFF">
		<tr>
			<td>
			<!--#include file="bottom.htm" -->
			</td>
		</tr>
	</table>
</div>

</body>

</html>