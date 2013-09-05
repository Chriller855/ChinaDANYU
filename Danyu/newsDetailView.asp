<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<html>
<title>公司新闻</title>
<link rel="stylesheet" type="text/css" href="css/base.css">
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
			<td><img border="0" src="images/b02.jpg" width="940" height="184"></td>
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
<div id="menu"><a href="news01.asp" id="m1"></a> <a href="news02.asp" id="m2"></a>
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
        <%
sql="select top 1 * from article where id="& clng(request("id"))
'response.write sql
Set Rs=Server.CreateObject("Adodb.Recordset")
rs.open sql,conn,1,1
response.write "<BR><center><font color=#6b6b6b style='font-size:13px;'><strong>"
response.write rs("news_title")
response.write "</strong></font><br>"
response.write "<font color=#8f8c8c>"&year(rs("news_uptime"))&"."&month(rs("news_uptime"))&"."&day(rs("news_uptime"))&"</font></center><BR><BR>"
response.write "<font color=#6b6b6b>"
response.write rs("news_content")
response.write "</font>"
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
                        </td>
					</tr>
                    <tr><td><p align="right" style="margin-right: 8"><font color="#6a6a6a">[ </font><a href="#" onClick="history.go(-1)"><font color="#c91213">返回</font></a><font color="#6a6a6a"> ]</font>   
    </td></tr>
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
			<div align="center">
	<table border="0" cellpadding="0" cellspacing="0" width="940" height="66" id="table1">
		<tr>
			<td width="119" align="center">
			<p style="margin-left: 5px">
			<img border="0" src="images/logo-b.png"></td>
			<td align="center">
			<img border="0" src="images/line01.gif" width="11" height="64"></td>
			<td width="390">&nbsp;&nbsp;&nbsp;
			<a target="_self" href="about01.asp">关于丹育</a>&nbsp;&nbsp;
			<font color="#CCCCCC">|</font><font color="#BBBBBB"> </font>&nbsp; <a href='javascript:location.reload()'>北欧农庄</a>&nbsp;&nbsp; 
			<font color="#CCCCCC">|</font>&nbsp;&nbsp;
			<a target="_self" href="news01.asp">公司新闻</a>&nbsp;&nbsp;
			<font color="#CCCCCC">|</font>&nbsp;&nbsp;
			<a target="_self" href="contact01.asp">联系我们</a></td>
			<td align="right">
			<map name="FPMap0">
			<area target="_blank" href="http://www.00ok.cc" shape="rect" coords="262, 19, 359, 40">
			</map>
			<img border="0" src="images/00ok.gif" width="374" height="45" usemap="#FPMap0"></td>
		</tr>
	</table>
</div>
			</td>
		</tr>
	</table>
</div>

</body>

</html>