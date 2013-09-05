<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!-- #include file="inc_config.asp" -->
<%
'权限检查
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<title><%=countname%>－数据清理－第一步－准备</title>
<link rel="stylesheet" type="text/css" href="../../index.css">
</head>
<body topmargin=5 rightmargin=0 leftmargin=0 vlink=#000000>
<table width="80%" cellspacing="0" align="center" cellpadding="0" border="0">
<form action="clear_step2.asp">
  <tr height="30">
    <td width="1" class="backs"></td>
    <td>
    
		数据清理－第一步&nbsp; □ 准备 ∷∷∷<br>

<table width="90%" align=center><tr><td>
<p class="p1">清理数据的步骤是这样的，首先系统按照你的要求清除主访问信息库中相应的记录，然后提示你使用数据库压缩工具对库文件进行压缩。
<p class="p1">1、在对数据库进行清理之前，请你确认你已经将你要清理的数据备份，并且你已经确认你不需要再在统计报告和自定义统计中使用这些信息了，因为这种清理是无法恢复的。
<p class="p1">2、如果您的服务器上没有安装FSO组件，程序就无法用压缩过的数据库覆盖原来的数据库，也无法删除临时的文件。你就必须准备好FTP软件等可以控制您的系统的其他办法，在清理完成后完成文件的更名。（<%
if IsObjInstalled("Scripting.FileSystemObject") then
%>您的空间支持FSO。<%
else%><font color=red>你的空间不支持FSO！<%end if%>）
<p class="p1"><font color=red>注意：</font>数据一旦从主访问记录库中清理出去，就无法在任何统计报告中看到这些记录，就是说不仅仅是在“详细记录”中看不到，而是在所有的统计报告中都看不到，自定义统计功能也无法查询到被清理掉的部分。使用“后备库查看器”可以查看被备份过的数据，但是功能很有限，所以应该在保证程序正常运行的情况下，保留尽量多的详细记录。（目前后备库查看器尚未完成）</p>
<p class="p1">准备好以后，请输入要清理的数据的条件并点击开始。
<%if Request("offtime")<>"" then%>
<p class="p1">清理 <input name="offyear" class="input" size="5" value="<%=year(Request("offtime"))%>"> 年
		<input name="offmonth" class="input" size="3" value="<%=month(Request("offtime"))%>"> 月
		<input name="offday" class="input" size="3" value="<%=day(Request("offtime"))%>"> 日前的数据。<INPUT type="checkbox" name=force value=1>是否强制清理
<%else%>
<p class="p1">清理 <input name="offyear" class="input" size="5"> 年
		<input name="offmonth" class="input" size="3"> 月
		<input name="offday" class="input" size="3"> 日前的数据。<INPUT type="checkbox" name=force value=1>是否强制清理
<%end if%>
<p class="p1">注意：当选择了“强制清理”时，没有备份过的数据也将被清理，请慎用。建议此选项仅在数据库已经出错无法备份的情况下使用。
<p class="p1" align="right"><a href='javascript:document.forms[0].submit();'>下一步 清理数据 开始</a> <input type="submit" name="bakgo" value="GO"><font style="font-size:16px">&nbsp;</font>&nbsp;

</td></tr></table>

	</td>
  </tr>
</form>
</table>
</body>
</html>
<%
'组件检查函数
Function IsObjInstalled(strClassString)
On Error Resume Next
IsObjInstalled = False
Err = 0
Dim xTestObj
Set xTestObj = Server.CreateObject(strClassString)
If 0 = Err Then IsObjInstalled = True
Set xTestObj = Nothing
Err = 0
End Function
%>
