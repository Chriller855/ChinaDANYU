<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!-- #include file="inc_config.asp" -->
<%
'权限检查

offtime	=request("offtime")
force	=Request("force")


if force="1" then
	conn.execute("delete * from view where vtime<datevalue('" & offtime & "')")
else
	conn.Execute("delete * from view where vtime<datevalue('" & offtime & "') and bakdays=1 and bakstats=1 and bakpage=1")
end if

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<title><%=countname%>－数据清理－第三步－清理</title>
<link rel="stylesheet" type="text/css" href="../../index.css">
</head>
<body topmargin=5 rightmargin=0 leftmargin=0 vlink=#000000>
<table width="80%" cellspacing="0" align="center" cellpadding="0" border="0">
  <tr height="30">
    <td width="1" class="backs"></td>
    <td class="backq">
    
		数据清理－第三步&nbsp; □ 清理指定记录 ∷∷∷<br>

<table width="90%" align=center><tr>
          <td>
<p class="p1">已经完成了数据清理。建议对数据库进行压缩，否则可能原来占用的空间并不能释放。
<p class="p1">数据压缩需要您的主机有足够的剩余空间，剩余空间的大小应该足够放置压缩后的数据库文件，一般应该是1～5M。如果您的主机没有足够的空间，就必须先清理一下主机上的文件，留出空间。
<p class="p1" align="right"><a href='clear_step4.asp'>下一步 压缩数据库</a>  <font style="font-size:16px">&nbsp;</font>&nbsp;
</td></tr></table>

	</td>
  </tr>
</table>
</body>
</html>
