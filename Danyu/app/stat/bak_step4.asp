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
<title><%=countname%>－数据备份－第四步－备份页面信息</title>
<link rel="stylesheet" type="text/css" href="../../index.css">
</head>
<body topmargin=5 rightmargin=0 leftmargin=0 vlink=#000000>
<table width="80%" cellspacing="0" align="center" cellpadding="0" border="0">
  </tr>
  <tr height="30">
    <td width="1" class="backs"></td>
    <td class="backq">
    
		数据备份－第四步&nbsp; □ 备份页面信息 ∷∷∷<br>

<table width="90%" align=center><tr>
          <td>
<p class="p1">系统正在备份页面信息，这些信息包括来访页面（连接到被统计页面的网页）和被访问页面的访问数据。
<p class="p1">请注意下面的红色方框，当方框中出现“页面信息备份完成。”时，点击框中的相应连接可继续。如果5分钟后仍然没有显示上述字样，或者显示“该页无法显示”，则请点击下面的“继续”按钮，或者刷新本页。
<p class="p1"><A href="bak_step4_in.asp?offtime=<%=offtime%>" target=bak_step4_in>[继续]</a>
<p><table width="400" border=1 bordercolor=red align=center><tr><td>
<IFRAME src="bak_step4_in.asp?offtime=<%=offtime%>" name="bak_step4_in" width="400" height="120" marginwidth="0" marginheight="0" frameborder="0" scrolling="no"></IFRAME>
</td></tr></table>
<p class="p1" align="right"><FONT 
            color=gray><U>完成</U></FONT> <input type="submit" name="bakgo" class="backc2" value=" GO"><font style="FONT-SIZE: 16px">&nbsp;</font>&nbsp;</p>

</td></tr></table>

	</td>
  </tr>
</table>
</body>
</html>
