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
<title><%=countname%>�����ݱ��ݣ����Ĳ�������ҳ����Ϣ</title>
<link rel="stylesheet" type="text/css" href="../../index.css">
</head>
<body topmargin=5 rightmargin=0 leftmargin=0 vlink=#000000>
<table width="80%" cellspacing="0" align="center" cellpadding="0" border="0">
  </tr>
  <tr height="30">
    <td width="1" class="backs"></td>
    <td class="backq">
    
		���ݱ��ݣ����Ĳ�&nbsp; �� ����ҳ����Ϣ �ˡˡ�<br>

<table width="90%" align=center><tr>
          <td>
<p class="p1">ϵͳ���ڱ���ҳ����Ϣ����Щ��Ϣ��������ҳ�棨���ӵ���ͳ��ҳ�����ҳ���ͱ�����ҳ��ķ������ݡ�
<p class="p1">��ע������ĺ�ɫ���򣬵������г��֡�ҳ����Ϣ������ɡ���ʱ��������е���Ӧ���ӿɼ��������5���Ӻ���Ȼû����ʾ����������������ʾ����ҳ�޷���ʾ��������������ġ���������ť������ˢ�±�ҳ��
<p class="p1"><A href="bak_step4_in.asp?offtime=<%=offtime%>" target=bak_step4_in>[����]</a>
<p><table width="400" border=1 bordercolor=red align=center><tr><td>
<IFRAME src="bak_step4_in.asp?offtime=<%=offtime%>" name="bak_step4_in" width="400" height="120" marginwidth="0" marginheight="0" frameborder="0" scrolling="no"></IFRAME>
</td></tr></table>
<p class="p1" align="right"><FONT 
            color=gray><U>���</U></FONT> <input type="submit" name="bakgo" class="backc2" value=" GO"><font style="FONT-SIZE: 16px">&nbsp;</font>&nbsp;</p>

</td></tr></table>

	</td>
  </tr>
</table>
</body>
</html>
