<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!-- #include file="inc_config.asp" -->
<%
'Ȩ�޼��

'��ȡ��������
offyear		=trim(Request("offyear"))
offmonth	=trim(Request("offmonth"))
offday		=trim(Request("offday"))

offtime=offyear & "-" & offmonth & "-" & offday
if (not isdate(offtime)) then Response.End()'Response.Redirect "help.asp?error=����ȷ��дҪ�������ݵĽ�ֹ���ڡ�"

force	=Request("force")
if force<>"1" then Response.Redirect "clear_step3.asp?offtime=" & offtime
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<title><%=countname%>�����������ڶ�����ȷ��</title>
<link rel="stylesheet" type="text/css" href="../../index.css">
</head>
<body topmargin=5 rightmargin=0 leftmargin=0 vlink=#000000>
<br>
<table width="80%" cellspacing="0" align="center" cellpadding="0" border="0">
  <tr height="30">
    <td width="1"></td>
    <td>
    
		���������ڶ���&nbsp; �� ǿ������ȷ�� �ˡˡ�<br>

<table width="90%" align=center><tr>
          <td>
<p class="p1">��ѡ���ˡ�ǿ������ѡ�δ�����ݹ��������б��������Σ�ա�����������ݿ�����Ӵ�����±���ʱ���ڴ�������󣬾Ͳ�Ҫʹ��ǿ�������ߡ�
<p class="p1">����Ĵ���ǿ���������ݿ���
<p class="p1" align="right"><a href='clear_step3.asp?offtime=<%=offtime%>'>��������(��ǿ��)</a>  <font style="font-size:16px">&nbsp;</font>&nbsp;
<a href='clear_step3.asp?offtime=<%=offtime%>&force=1'>��������(ǿ��)</a>  <font style="font-size:16px">&nbsp;</font>&nbsp;
</td></tr></table>

	</td>
  </tr>
</table>
</body>
</html>
