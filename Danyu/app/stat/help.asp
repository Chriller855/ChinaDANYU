<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!-- #include file="inc_config.asp" -->
<%
'从浏览器获取参数
id=Request("id")
error=Request("error")
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<title><%=countname%>－帮助信息</title>
<link rel="stylesheet" type="text/css" href="../../index.css">
</head>
<body topmargin=5 rightmargin=0 leftmargin=0 vlink=#000000>
<!--#include file="inc_top.asp"-->
<br>
<table width="500" cellspacing="0" align="center" cellpadding="0" border="0">
  <tr><td colspan="3"><img src="images/photoup.gif"></td></tr>
  <tr height="30">
    <td width="1" class="backs"><img src="images/touming.gif" width="1" height="1"></td>
    <td width="498" class="backq">
		&nbsp; <img src="images/tb_help.gif" align=absmiddle> <font style="font-size:16px">&nbsp;</font>帮 助 信 息<br>
	<%if error <> "" then Response.Write "<p class='p1'><font color='red'><img src='images/tb_error.gif' align=absmiddle><font style='font-size:15px'>&nbsp;</font>错误 " & error & "</font>"%>
<%select case id
case ""%>
	<li style="	margin-left: 50; margin-top: 15;margin-bottom: 5"><a href="help.asp?id=001">自定义统计条件应该怎样填写呢？</a>
	<li style="	margin-left: 50; margin-top: 5;	margin-bottom: 5"><a href="help.asp?id=002">如何保存自定义检索分析的检索条件？</a>
	<li style="	margin-left: 50; margin-top: 5;	margin-bottom: 5"><a href="help.asp?id=003">如何修改或删除已经保存的检索条件？</a>
	<li style="	margin-left: 50; margin-top: 5;	margin-bottom: 5"><a href="help.asp?id=004">访客和管理员分别具有哪些权限？</a>
	<li style="	margin-left: 50; margin-top: 5;	margin-bottom: 5"><a href="help.asp?id=007">为什么系统提示我“该系统尚未启用”？</a>
	<li style="	margin-left: 50; margin-top: 5;	margin-bottom: 5"><a href="help.asp?id=005">如何获得本系统的最新版本？</a>
	<li style="	margin-left: 50; margin-top: 5;	margin-bottom: 5"><a href="help.asp?id=006">本系统的版权声明。</a>
	
<%case "001"%>
	<p class="p1"><font class="fonts"><b>自定义统计条件应该怎样填写呢？</b></font>
	<p class="p1">起止日期：年、月、日都必须填写，而且都必须是数字。但是可以只填起始日期，这样表示从填写的起始日期至现在这个时间段，同样也可以只填写截止日期，这样表示从开始统计那天至所填写的这个日期，注意截止日期必须在开通统计日期之后。
	<p class="p1">IP地址和来自地区：这两项都是支持模糊查询的，例如在IP地址项中输入“61.”，则可查到IP地址是“61.163.233.10”或者“202.61.33.25”这样的访问记录的统计分析。
	<p class="p1">操作系统和浏览器：这两项可以从列表中选择，也可以直接在后边文本框中输入，也支持模糊查询，比如在文本框中输入win，则可以查到WIN 2K、WIN XP、WIN 9X的全部记录的统计分析。
	<p class="p1">来自和被访网页：这两个选项可以让您查到特定来路或目标的访问记录的统计分析，支持模糊查询。
	<p class="p1">输出类型：可以选择要查看的分析的类型，可多选，但要注意您选择的越多，需要等待的时间就越长，当然，您也不能一个都不选。

<%case "002"%>
	<p class="p1"><font class="fonts"><b>如何保存自定义检索分析的检索条件？</b></font>
	<p class="p1">在进行了自定义统计的检索之后，在该页面底部有一个名为“保存本次检索条件”的项目框，该框中的选项允许你为要保存的条件取名并加入简介。
	<p class="p1">保存前必须为该条件取一个名字，否则日后将无法使用您保存的检索条件。
	<p class="p1">最好再为要保存的检索条件写一个简介，因为名字写的太长会影响页面美观的，要细致得看懂名字所代表的意思，写个简介是必要的。
	<p class="p1">在名字的输入框后面有两个单选框，可以让您指明当和已有的检索项目重名是覆盖掉原有的呢还是让系统提示重新取名字，除非您本来就是为了修改已有项目并确认名字没写错的话，否则一定保持它选择“同名时提问”，以确保已有数据的安全。
	<p class="p1">只有管理员才可以进行上述操作。
	
<%case "003"%>
	<p class="p1"><font class="fonts"><b>如何修改或删除已经保存的检索条件？</b></font>
	<p class="p1">［修改］您只需要按照新的检索条件进行检索，在检索结果页面的底部的保存项中使用原有的名字保存，并选择“同名时覆盖”，保存后就修改了原来保存的项目。
	<p class="p1">［删除］在“自定义统计”页面的自定义检索条件列表中点击“删除”连接，进入删除页面，系统提示是否真的要删除，点击确定即可。
	<p class="p1">只有管理员才可以进行上述操作。

<%case "004"%>
	<p class="p1"><font class="fonts"><b>访客和管理员分别具有哪些权限？</b></font>
	<p class="p1">阿江酷站访问统计对访问权限实行等级管理，管理员始终具有 6 级的权力，即拥有所有权力，非管理员（访客）所拥有的权限由管理员在安装时在配置文件中设置。等级划分的办法是：</p>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="320" align=center>
	<tr height="18" align=center class=backzq><td width="110">等 级</td><td width="30">0</td><td width="30">1</td><td width="30">2</td><td width="30">3</td><td width="30">4</td><td width="30">5</td><td width="30">6</td></tr>
	<tr height="18" align=center><td align=left>&nbsp;查看总体数据</td><td></td><td>√</td><td>√</td><td>√</td><td>√</td><td>√</td><td>√</td></tr> 
	<tr height="18" align=center class=backzq><td align=left>&nbsp;查看缺省统计数据</td><td></td><td></td><td>√</td><td>√</td><td>√</td><td>√</td><td>√</td></tr> 
	<tr height="18" align=center><td align=left>&nbsp;查看详细记录</td><td></td></td><td></td><td><td>√</td><td>√</td><td>√</td><td>√</td></tr> 
	<tr height="18" align=center class=backzq><td align=left>&nbsp;查看自定义记录</td><td></td><td></td><td></td><td></td><td>√</td><td>√</td><td>√</td></tr> 
	<tr height="18" align=center><td align=left>&nbsp;自定义条件统计</td><td></td><td></td><td></td><td></td><td></td><td>√</td><td>√</td></tr> 
	<tr height="18" align=center class=backzq><td align=left>&nbsp;管理自定义库</td><td></td><td></td><td></td><td></td><td></td><td></td><td>√</td></tr> 
</table>
	<p class="p1">如果访客等级被设置为 0，则访客将没有任何权力，如果访客等级被设置为 6，则访客将具有和管理员相等的权力。默认的访客权限等级为 4。
	<p class="p1">当前访客的权限等级为：<%=whatcan%>

<%case "005"%>
	<p class="p1"><font class="fonts"><b>如何获得本系统的最新版本？</b></font>
	<p class="p1">请您关注我的主页 <a href="http://www.ajiang.net" target="_blank">http://www.ajiang.net</a> 我会在推出新版的第一时间在主页上发布。您也将从阿江的主页上获得阿江编写的其他程序。
	<p class="p1">请务必注意保留阿江的版权信息，我花费很多精力开发各种程序给大家免费使用，会有一种成就感，我不想别人窃取这种荣耀。
	<p class="p1">最后感谢您使用阿江编写的程序，希望可以得到您的建议和继续支持。

<%case "006"%>
	<p class="p1"><font class="fonts"><b>本系统的版权信息</b></font>
	<p class="p1">本软体为共享软体(shareware)提供个人网站免费使用，作者阿江(ajiang.net)保留全部权力，受相关法律和国际公约保护，请勿非法修改、转载、散播，或用于其他赢利行为，并请勿删除版权声明。
<p class="p1">希望您在启用本系统后到我的网站订阅本软件的更新通知邮件，本系统大约每周更新一次。
<p class="p1">如果您在使用中遇到了问题，请到我的网站提问。
<%case "007"%>
	<p class="p1"><font class="fonts"><b>为什么系统提示我“该系统尚未启用”？</b></font>
	<p class="p1">这是因为您统计结果还是0，这可能是因为您还没有在被统计页面放上统计代码，或者放上了以后被统计页面还没有被访问过。统计系统的嵌入代码是：
	<p class="p1">&lt;script src="<%="http://" & Request.ServerVariables("http_host") & finddir(Request.ServerVariables("SCRIPT_NAME"))%>mystat.asp?style=icon"&gt;&lt;/script&gt;
	<p class="p1">有关嵌入代码的写法请参考本系统的安装说明。
	<p class="p1">如果您确认您已经正确放上来以上代码并访问过，则请检查您的服务器是否允许写入。
	
<%case else%>
<p class="p1">目前还没有与此相关的帮助信息。

<%end select%>
<p class="p1" align="right"><a href='javascript:history.back()'>继续</a> <a href='javascript:history.back()'><img src="images/arbutton.gif" align="absmiddle" border="0"></a> <font style="font-size:16px">&nbsp;</font>&nbsp;
	</td>
    <td width="1" class="backs"><img src="images/touming.gif" width="1" height="1"></td>
  </tr>
  <tr><td colspan="4"><img src="images/photodown.gif"></td></tr>
</table>
<br>
<!--#include file="inc_bottom.asp"-->
</body>
</html>
<%
Function finddir(filepath)
	finddir=""
	for i=1 to len(filepath)
	if left(right(filepath,i),1)="/" or left(right(filepath,i),1)="\" then
	  abc=i
	  exit for
	end if
	next
	if abc <> 1 then
	finddir=left(filepath,len(filepath)-abc+1)
	end if
end Function
%>