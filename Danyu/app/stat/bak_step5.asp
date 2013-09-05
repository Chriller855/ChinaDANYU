<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!-- #include file="inc_config.asp" -->
<%
'权限检查

offtime=Request("offtime")
if (not isdate(offtime)) then Response.End()'Response.Redirect "help.asp?error=请正确填写要备份数据的截止日期。"

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<title><%=countname%>－数据备份－第五步－完成</title>
<link rel="stylesheet" type="text/css" href="../../index.css">
</head>
<body topmargin=5 rightmargin=0 leftmargin=0 vlink=#000000>
<table width="80%" cellspacing="0" align="center" cellpadding="0" border="0">
  <tr height="30">
    <td width="1" class="backs"></td>
    <td>
    
		数据备份－第五步&nbsp; □ 完成 ∷∷∷<br>

<table width="90%" align=center><tr>
          <td>
<p class="p1">您成功完成了对<%=year(offtime)%>年<%=month(offtime)%>月<%=day(offtime)%>日前的数据的备份。
<p class="p1">如果您暂时还不想清理数据，请点击关闭窗口或者随便到你想去的网页去逛。如果您要清理数据（也许您感到您的数据库实在是太大了），请点击下面的清理数据连接。
<p align="right"><a href="" onclick="window.close()">关闭窗口</a> &nbsp; 
<a href='clear_step1.asp?offtime=<%=offtime%>'>清理数据</a>  <font style="font-size:16px">&nbsp;</font>&nbsp;

</td></tr></table>

	</td>
  </tr>
</table>
<br>
</body>
</html>
