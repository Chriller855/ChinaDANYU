<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!-- #include file="inc_config.asp" -->
<%
'Ȩ�޼��

offtime=Request("offtime")
if (not isdate(offtime)) then Response.End()'Response.Redirect "help.asp?error=����ȷ��дҪ�������ݵĽ�ֹ���ڡ�"

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<title><%=countname%>�����ݱ��ݣ����岽�����</title>
<link rel="stylesheet" type="text/css" href="../../index.css">
</head>
<body topmargin=5 rightmargin=0 leftmargin=0 vlink=#000000>
<table width="80%" cellspacing="0" align="center" cellpadding="0" border="0">
  <tr height="30">
    <td width="1" class="backs"></td>
    <td>
    
		���ݱ��ݣ����岽&nbsp; �� ��� �ˡˡ�<br>

<table width="90%" align=center><tr>
          <td>
<p class="p1">���ɹ�����˶�<%=year(offtime)%>��<%=month(offtime)%>��<%=day(offtime)%>��ǰ�����ݵı��ݡ�
<p class="p1">�������ʱ�������������ݣ������رմ��ڻ�����㵽����ȥ����ҳȥ�䡣�����Ҫ�������ݣ�Ҳ�����е��������ݿ�ʵ����̫���ˣ�����������������������ӡ�
<p align="right"><a href="" onclick="window.close()">�رմ���</a> &nbsp; 
<a href='clear_step1.asp?offtime=<%=offtime%>'>��������</a>  <font style="font-size:16px">&nbsp;</font>&nbsp;

</td></tr></table>

	</td>
  </tr>
</table>
<br>
</body>
</html>
