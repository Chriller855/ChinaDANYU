<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!-- #include file="inc_config.asp" -->
<%
'权限检查

'获取备份条件
offyear		=trim(Request("offyear"))
offmonth	=trim(Request("offmonth"))
offday		=trim(Request("offday"))

offtime=offyear & "-" & offmonth & "-" & offday
if (not isdate(offtime)) then Response.End()'Response.Redirect "help.asp?error=请正确填写要备份数据的截止日期。"

force	=Request("force")
if force<>"1" then Response.Redirect "clear_step3.asp?offtime=" & offtime
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<title><%=countname%>－数据清理－第二步－确认</title>
<link rel="stylesheet" type="text/css" href="../../index.css">
</head>
<body topmargin=5 rightmargin=0 leftmargin=0 vlink=#000000>
<br>
<table width="80%" cellspacing="0" align="center" cellpadding="0" border="0">
  <tr height="30">
    <td width="1"></td>
    <td>
    
		数据清理－第二步&nbsp; □ 强制清理确认 ∷∷∷<br>

<table width="90%" align=center><tr>
          <td>
<p class="p1">你选择了“强制清理”选项，未被备份过的数据有被清除掉的危险。建议除非数据库过于庞大而导致备份时的内存溢出错误，就不要使用强制清理工具。
<p class="p1">您真的打算强制清理数据库吗？
<p class="p1" align="right"><a href='clear_step3.asp?offtime=<%=offtime%>'>清理数据(不强制)</a>  <font style="font-size:16px">&nbsp;</font>&nbsp;
<a href='clear_step3.asp?offtime=<%=offtime%>&force=1'>清理数据(强制)</a>  <font style="font-size:16px">&nbsp;</font>&nbsp;
</td></tr></table>

	</td>
  </tr>
</table>
</body>
</html>
