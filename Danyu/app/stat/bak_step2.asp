<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!-- #include file="inc_config.asp" -->
<%
'��ȡ��������
offyear		=trim(Request("offyear"))
offmonth	=trim(Request("offmonth"))
offday		=trim(Request("offday"))

offtime=offyear & "-" & offmonth & "-" & offday
'if (not isdate(offtime)) then Response.Redirect "help.asp?error=����ȷ��дҪ�������ݵĽ�ֹ���ڡ�"
if (not isdate(offtime)) then
Response.write "����ȷ��дҪ�������ݵĽ�ֹ���ڡ�"
response.End()
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<title><%=countname%>�����ݱ��ݣ��ڶ���������ÿ������</title>
<link rel="stylesheet" type="text/css" href="../../index.css">
</head>
<body topmargin=5 rightmargin=0 leftmargin=0 vlink=#000000>
<br>
<table width="80%" cellspacing="0" align="center" cellpadding="0" border="0">
  <tr height="30">
    <td width="1" class="backs"></td>
    <td class="backq">
    
		���ݱ��ݣ��ڶ���&nbsp; �� ����ÿ������ �ˡˡ�<br>

<table width="90%" align=center><tr>
          <td>
<p class="p1">ϵͳ���ڱ���ÿ�յķ������ݣ���Щ���ݰ���ÿ�յķ�������ÿ�շ����ߵ�IP��ַ����Ŀ��
<p class="p1">��ע������ĺ�ɫ���򣬵������г��֡�ÿ�����ݱ�����ɡ���ʱ��������еġ���һ�������ӿɼ��������3���Ӻ�����ĺ�ɫ��������Ȼû����ʾ����������������ʾ����ҳ�޷���ʾ��������������ġ���������ť������ˢ�±�ҳ��
<p class="p1"><A href="bak_step2_in.asp?offtime=<%=offtime%>" target=bak_step2_in>[����]</a>
<p><table width="400" border=1 bordercolor=red align=center><tr><td>
<IFRAME src="bak_step2_in.asp?offtime=<%=offtime%>" name="bak_step2_in" width="400" height="120" marginwidth="0" marginheight="0" frameborder="0" scrolling="no"></IFRAME>
</td></tr></table>
<p class="p1" align="right"><FONT 
            color=gray><U>��һ�� ���ݿͻ�����Ϣ ��ʼ</U></FONT> <input type="submit" name="bakgo" value="GO"><font style="FONT-SIZE: 16px">&nbsp;</font>&nbsp;</p>

</td></tr></table>

	</td>
  </tr>
</table>
<br>
</body>
</html>
