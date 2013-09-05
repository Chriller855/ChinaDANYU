<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!-- #include file="inc_config.asp" -->
<%
'获取备份条件
offyear		=trim(Request("offyear"))
offmonth	=trim(Request("offmonth"))
offday		=trim(Request("offday"))

offtime=offyear & "-" & offmonth & "-" & offday
'if (not isdate(offtime)) then Response.Redirect "help.asp?error=请正确填写要备份数据的截止日期。"
if (not isdate(offtime)) then
Response.write "请正确填写要备份数据的截止日期。"
response.End()
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<title><%=countname%>－数据备份－第二步－备份每日数据</title>
<link rel="stylesheet" type="text/css" href="../../index.css">
</head>
<body topmargin=5 rightmargin=0 leftmargin=0 vlink=#000000>
<br>
<table width="80%" cellspacing="0" align="center" cellpadding="0" border="0">
  <tr height="30">
    <td width="1" class="backs"></td>
    <td class="backq">
    
		数据备份－第二步&nbsp; □ 备份每日数据 ∷∷∷<br>

<table width="90%" align=center><tr>
          <td>
<p class="p1">系统正在备份每日的访问数据，这些数据包括每日的访问量和每日访问者的IP地址的数目。
<p class="p1">请注意下面的红色方框，当方框中出现“每日数据备份完成。”时，点击框中的“下一步”连接可继续。如果3分钟后下面的红色方框中仍然没有显示上述字样，或者显示“该页无法显示”，则请点击下面的“继续”按钮，或者刷新本页。
<p class="p1"><A href="bak_step2_in.asp?offtime=<%=offtime%>" target=bak_step2_in>[继续]</a>
<p><table width="400" border=1 bordercolor=red align=center><tr><td>
<IFRAME src="bak_step2_in.asp?offtime=<%=offtime%>" name="bak_step2_in" width="400" height="120" marginwidth="0" marginheight="0" frameborder="0" scrolling="no"></IFRAME>
</td></tr></table>
<p class="p1" align="right"><FONT 
            color=gray><U>下一步 备份客户端信息 开始</U></FONT> <input type="submit" name="bakgo" value="GO"><font style="FONT-SIZE: 16px">&nbsp;</font>&nbsp;</p>

</td></tr></table>

	</td>
  </tr>
</table>
<br>
</body>
</html>
