<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!-- #include file="inc_config.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<title><%=countname%>－数据备份－第一步</title>
<link rel="stylesheet" type="text/css" href="../../index.css">
</head>
<body topmargin=5 rightmargin=0 leftmargin=0 vlink=#000000>
<table width="80%" cellspacing="0" align="center" cellpadding="0" border="0">
<form action="bak_step2.asp">
  <tr height="30">
    <td width="1" class="backs"></td>
    <td class="backq">
    
		数据备份－第一步&nbsp; □ 准备 ∷∷∷<br>

<table width="90%" align=center><tr><td>
<p>备份数据要占用大量系统资源，而且可能会因为中断而导致出错，所以请选择网站访问量较小的时段进行本操作。
<p>备份过的记录仍然保存在主数据库中，可以在统计报告和自定义统计中看到，只有使用了“数据清理”工具之后才会从主库中删除。
<p>准备好以后，请输入要备份的数据的条件并点击开始。系统对备份过的数据做了标记，所以不会重复备份。
<p>如果执行中出错，比如提示脚本超时（script timeout），就请点击浏览器上的刷新（或reload）按钮刷新页面，程序会自动继续执行，你也可以不去管它，下次备份时会自动备份这次漏掉的数据。
<p>为了减轻程序负担，每次备份的数据量最好不要超过2万条，比如您的网站每月访问4万次，则请您在填写下面的条件时确定每次备份半个月。
<p>备份 <input name="offyear" class="input" size="5"> 年
		<input name="offmonth" class="input" size="3"> 月
		<input name="offday" class="input" size="3"> 日前的数据。
<a href='javascript:document.forms[0].submit();'>下一步 备份每日数据 开始</a> <input type="submit" name="bakgo" value=" GO">
&nbsp;

</td></tr></table>

	</td>
  </tr>
</form>
</table>
<br>
<!--#include file="inc_bottom.asp"-->
</body>
</html>
