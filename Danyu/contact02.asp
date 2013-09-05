<!-- #include file="app/notepad1/conn.asp" -->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>在线咨询</title>
<!--<link rel="stylesheet" type="text/css" href="css/base.css">-->
<link rel="stylesheet" type="text/css" href="css/atsunan.css">
<link rel="stylesheet" type="text/css" href="css/contact.css">
<link rel="stylesheet" type="text/css" href="index.css">
<script type="text/javascript" src="Scripts/prototype.js" ></script>
<script type="text/javascript" src="Scripts/contact_current.js"></script>
<script src="Scripts/AC_RunActiveContent.js" type="text/javascript"></script>
    <style type="text/css">
        .auto-style1 {
            height: 27px;
        }
    </style>
</head>

<body topmargin="0" leftmargin="0" bgcolor="#F5F5F5">

<div align="center">
	<table border="0" cellpadding="0" cellspacing="0" id="table1" width="940" height="80" bgcolor="#FFFFFF">
		<tr>
			<td>
			  <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="940" height="80">
          <param name="movie" value="flash/menu.swf?id=7">
		  <param name="wmode" value="transparent">
          <param name="quality" value="high">
          <embed src="flash/menu.swf?id=7" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="940" height="80">
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
<div id="menu"><a href="contact01.asp" id="m1"></a> <a href="contact02.asp" id="m2"></a>
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
						<img border="0" src="images/tit0702.gif" width="691" height="47"></td>
					</tr>
					<tr>
						<td width="670">
                        <table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td  align="left" valign="top" height="55" background="images/line03.gif">
    <p align="left"><font color="#6a6a6a">[&nbsp;</font><a href="javascript:;" onClick="document.body.scrollTop=10000;"><strong><font color="#d21213">我要留言</font></strong></a><font color="#6a6a6a">&nbsp;]</font></td>

  </tr>
  <tr> 
    <td> <%
if request("submit")="发 送" then
	if datediff("s",session("notePadUpdateTime"),now())<180 then
		'response.write "<div align=center><font color=#c80000 style='font-size:14px;'>请勿重复提交! 或过180秒后重新留言！！</font></div>"
		response.write "<script language=javascript>"
		response.write "alert('请勿重复提交! 或过180秒后重新留言！！');"
		response.write "location.href='"&request.ServerVariables("PATH_INFO")&"';"
		response.write "</script>"
		response.end
		'conn.close
		'set conn=nothing
	else
		if request("saveCookies")="on" then
		response.Cookies("YoungXsTianYuan").Expires=Date+365
		response.Cookies("YoungXsTianYuan")("name")=request("name")
		response.Cookies("YoungXsTianYuan")("email")=request("email")
		response.Cookies("YoungXsTianYuan")("tel")=request("tel")
		response.Cookies("YoungXsTianYuan")("address")=request("address")
		else
		response.Cookies("YoungXsTianYuan")("name")=""
		response.Cookies("YoungXsTianYuan")("email")=""
		response.Cookies("YoungXsTianYuan")("tel")=""
		response.Cookies("YoungXsTianYuan")("address")=""
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
		rs("validFlag")=0
		rs("name")=server.HTMLEncode(request("name"))
		rs("email")=server.HTMLEncode(request("email"))
		rs("tel")=server.HTMLEncode(request("tel"))
		rs("address")=server.HTMLEncode(request("address"))
		rs("content")=server.HTMLEncode(Request("comment"))
		rs("dateandtime")=now()
		rs("ip")=request.ServerVariables("REMOTE_ADDR")
		rs.updatebatch
		rs.close
		session("notePadUpdateTime")=now()
		'response.Redirect(request.ServerVariables("PATH_INFO"))
		response.write "<script language=javascript>"
		response.write "alert('谢谢您的关注，我们会尽快答复您！');"
		response.write "location.href='"&request.ServerVariables("PATH_INFO")&"';"
		response.write "</script>"
		response.end
	end if
end if
if session("notePadUpdateTime")="" then session("notePadUpdateTime")=cdate("1900-01-01")
session("refererString")=request.ServerVariables("SCRIPT_NAME")&"?page="&Request("Page")
%> <table width="95%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="110" align="center" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td align="center"> <%
set rs1=server.createobject("adodb.recordset")
set rs2=server.createobject("adodb.recordset")
'sql="select * from book1 where id="&request("id") & " or subid="&request("id")
sql1="select * from book1 where subid=0 and validFlag<>0 order by dateandtime desc"
rs1.open sql1,conn,1,1
totalrec=rs1.recordcount
'response.write totalrec
if rs1.eof then
	response.write "<br>"
else
const maxPerPage=5
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
%> <%
	i=-1
   For iPage = 1 To maxPerPage
    if rs1.eof then exit for
	
	i=i+1

	'sql2="select * from book1 where id="&rs1("id")&" or subid="&rs1("id")&" order by subid asc ,id desc"
	sql2="select * from book1 where id="&rs1("id")&" or subid="&rs1("id")&" order by  subid asc ,dateAndTime asc"
	'response.write sql2&"<br>"
	rs2.open sql2,conn,1,1
	do while not rs2.eof
%> <table width=100% border=0 cellpadding=8 cellspacing=0 <%if i mod 2 = 0 then
				 ' response.write "bgcolor=#DCF1F6"
				  else
				 ' response.write "bgcolor=#EAF8FB"
				  end if%>>
                    <tr> 
                      <td> <%
if rs2("subid")=0 then	
	response.write "<font color=#000000 style='font-size:12px;'>"
else
	response.write "<strong><font color=#d21213 style='font-size:12px;'>"
end if
	if rs2("name")="" then
		response.write "NoName"
	else
		response.write rs2("name")
	end if
response.write "&nbsp;"
response.write year(rs2("dateandtime"))&"-"&month(rs2("dateandtime"))&"-"&day(rs2("dateandtime"))
response.write "</font></strong>"
response.write "<br><font "
if rs2("subid")=0 then
response.write "color=#6a6a6a"
else
response.write "color=#6a6a6a"
end if
response.write ">"
response.write replace(cstr(rs2("content")),vbCrLf,"<br>")
response.write "</font>"
%></td>
                    </tr>
                  </table>
                  <%
rs2.movenext
loop
rs2.close
rs1.movenext
response.write "<br><hr size=1 color=#cccccc>"
next
rs1.close
set rs1=nothing
set rs2=nothing
%> </td>
              </tr>
              <tr> 
                <td height="40" align="right"> <%				  
if totalrec<>0 then
response.write "<font color=#4E4E4E>Total <b>"&totalrec&"</b></font> "
	if currentpage-1 mod 10=0 then
		p=(currentpage-1) \ 10
	else
		p=(currentpage-1) \ 10
	end if
	dim pagelist,pagelistbit
	if currentPage=1 then
		response.write "<font face=webdings color=#555555>9</font>   "
	else
		response.write "<a href='?sortid="&sortid&"&page=1' title=首页><font face=webdings>9</font></a>   "
	end if
	if p*10>0 then response.write "<a href='?sortid="&sortid&"&page="&Cstr(p*10)&"' title=上十页><font face=webdings>7</font></a>   "
	response.write "<b>"
	for ii=p*10+1 to P*10+10
		if ii=currentPage then
			response.write "<font color=#BABABA>"+Cstr(ii)+"</font> "
		else
			response.write "<a href='?sortid="&sortid&"&page="&Cstr(ii)&"'>"+Cstr(ii)+"</a>   "
		end if
		if ii=n then exit for
		'p=p+1
	next
	response.write "</b>"
	if ii<n then response.write "<a href='?sortid="&sortid&"&page="&Cstr(ii)&"' title=下十页><font face=webdings>8</font></a>   "
	if currentPage=n then
		response.write "<font face=webdings color=#555555>:</font>   "
	else
		response.write "<a href='?sortid="&sortid&"&page="&Cstr(n)&"' title=尾页><font face=webdings>:</font></a>"
	end if
end if
%> <%
end if
%> </td>
              </tr>
            </table></td>
          <td width="2%" valign="top">　</td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td align="center"><div id=aBook></div>
    <table width="100%" border="0" cellspacing="0" cellpadding="3"><form action="<%=request.ServerVariables("PATH_INFO")%>" method="POST" style="padding:0; ">
          <td width="64" align=left><font color="#222222"> 姓 名：</font></td>
          <td width="610"> 
        <input type=text name=name size="25" value="<%=request.Cookies("YoungXsTianYuan")("name")%>" style="padding:0; "></td>
        </tr>
        <tr> 
          <td align=left><font color="#222222">Email：</font></td>
          <td> 
          <input type=text name=email size="25" value="<%=request.Cookies("YoungXsTianYuan")("email")%>" style="padding:0; "></td>
        </tr>
        <tr> 
          <td align=left class="auto-style1"> <font color="#222222">电 话：</font></td>
          <td class="auto-style1"> 
          <input name=tel type=text style="padding:0;" value="<%=request.Cookies("YoungXsTianYuan")("tel")%>" size="25"></td>
        </tr>
        <tr> 
          <td align=left> <p style="margin-top: 5"><font color="#222222">留 言：</font></td>
          <td> 　</td>
        </tr>
        <tr> 
          <td align="left" valign="top" colspan="2"> <p style="margin-left: 2"> 
              <textarea cols="100" rows=10 name=comment wrap="soft" style="padding:5px; background-color: #f5f5f5;"></textarea>
          </td>
        </tr>
        <tr> 
          <td colspan="2"> <input name="submit" type=submit value="发 送" style="background-color: #f5f5f5; color: #222222; border: 1 solid #bbbbbb"> &nbsp;&nbsp; 
            <input name="saveCookies" type="checkbox" class="noBorder" style="border: 0px;" value="on" checked> 
            <font color="#525252"> <span class="small">Save Cookies</span></font><br> 
            <br> </td><td width="-1"></form>
      </table></td>
  </tr>
</table>
                        </td>
					</tr>
					<tr>
						<td>&nbsp;</td>
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