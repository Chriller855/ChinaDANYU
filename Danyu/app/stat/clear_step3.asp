<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!-- #include file="inc_config.asp" -->
<%
'Ȩ�޼��

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
<title><%=countname%>����������������������</title>
<link rel="stylesheet" type="text/css" href="../../index.css">
</head>
<body topmargin=5 rightmargin=0 leftmargin=0 vlink=#000000>
<table width="80%" cellspacing="0" align="center" cellpadding="0" border="0">
  <tr height="30">
    <td width="1" class="backs"></td>
    <td class="backq">
    
		��������������&nbsp; �� ����ָ����¼ �ˡˡ�<br>

<table width="90%" align=center><tr>
          <td>
<p class="p1">�Ѿ����������������������ݿ����ѹ�����������ԭ��ռ�õĿռ䲢�����ͷš�
<p class="p1">����ѹ����Ҫ�����������㹻��ʣ��ռ䣬ʣ��ռ�Ĵ�СӦ���㹻����ѹ��������ݿ��ļ���һ��Ӧ����1��5M�������������û���㹻�Ŀռ䣬�ͱ���������һ�������ϵ��ļ��������ռ䡣
<p class="p1" align="right"><a href='clear_step4.asp'>��һ�� ѹ�����ݿ�</a>  <font style="font-size:16px">&nbsp;</font>&nbsp;
</td></tr></table>

	</td>
  </tr>
</table>
</body>
</html>
