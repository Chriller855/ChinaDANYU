<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!-- #include file="inc_config.asp" -->
<%
'���������ȡ����
id=Request("id")
error=Request("error")
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<title><%=countname%>��������Ϣ</title>
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
		&nbsp; <img src="images/tb_help.gif" align=absmiddle> <font style="font-size:16px">&nbsp;</font>�� �� �� Ϣ<br>
	<%if error <> "" then Response.Write "<p class='p1'><font color='red'><img src='images/tb_error.gif' align=absmiddle><font style='font-size:15px'>&nbsp;</font>���� " & error & "</font>"%>
<%select case id
case ""%>
	<li style="	margin-left: 50; margin-top: 15;margin-bottom: 5"><a href="help.asp?id=001">�Զ���ͳ������Ӧ��������д�أ�</a>
	<li style="	margin-left: 50; margin-top: 5;	margin-bottom: 5"><a href="help.asp?id=002">��α����Զ�����������ļ���������</a>
	<li style="	margin-left: 50; margin-top: 5;	margin-bottom: 5"><a href="help.asp?id=003">����޸Ļ�ɾ���Ѿ�����ļ���������</a>
	<li style="	margin-left: 50; margin-top: 5;	margin-bottom: 5"><a href="help.asp?id=004">�ÿͺ͹���Ա�ֱ������ЩȨ�ޣ�</a>
	<li style="	margin-left: 50; margin-top: 5;	margin-bottom: 5"><a href="help.asp?id=007">Ϊʲôϵͳ��ʾ�ҡ���ϵͳ��δ���á���</a>
	<li style="	margin-left: 50; margin-top: 5;	margin-bottom: 5"><a href="help.asp?id=005">��λ�ñ�ϵͳ�����°汾��</a>
	<li style="	margin-left: 50; margin-top: 5;	margin-bottom: 5"><a href="help.asp?id=006">��ϵͳ�İ�Ȩ������</a>
	
<%case "001"%>
	<p class="p1"><font class="fonts"><b>�Զ���ͳ������Ӧ��������д�أ�</b></font>
	<p class="p1">��ֹ���ڣ��ꡢ�¡��ն�������д�����Ҷ����������֡����ǿ���ֻ����ʼ���ڣ�������ʾ����д����ʼ�������������ʱ��Σ�ͬ��Ҳ����ֻ��д��ֹ���ڣ�������ʾ�ӿ�ʼͳ������������д��������ڣ�ע���ֹ���ڱ����ڿ�ͨͳ������֮��
	<p class="p1">IP��ַ�����Ե������������֧��ģ����ѯ�ģ�������IP��ַ�������롰61.������ɲ鵽IP��ַ�ǡ�61.163.233.10�����ߡ�202.61.33.25�������ķ��ʼ�¼��ͳ�Ʒ�����
	<p class="p1">����ϵͳ�����������������Դ��б���ѡ��Ҳ����ֱ���ں���ı��������룬Ҳ֧��ģ����ѯ���������ı���������win������Բ鵽WIN 2K��WIN XP��WIN 9X��ȫ����¼��ͳ�Ʒ�����
	<p class="p1">���Ժͱ�����ҳ��������ѡ����������鵽�ض���·��Ŀ��ķ��ʼ�¼��ͳ�Ʒ�����֧��ģ����ѯ��
	<p class="p1">������ͣ�����ѡ��Ҫ�鿴�ķ��������ͣ��ɶ�ѡ����Ҫע����ѡ���Խ�࣬��Ҫ�ȴ���ʱ���Խ������Ȼ����Ҳ����һ������ѡ��

<%case "002"%>
	<p class="p1"><font class="fonts"><b>��α����Զ�����������ļ���������</b></font>
	<p class="p1">�ڽ������Զ���ͳ�Ƶļ���֮���ڸ�ҳ��ײ���һ����Ϊ�����汾�μ�������������Ŀ�򣬸ÿ��е�ѡ��������ΪҪ���������ȡ���������顣
	<p class="p1">����ǰ����Ϊ������ȡһ�����֣������պ��޷�ʹ��������ļ���������
	<p class="p1">�����ΪҪ����ļ�������дһ����飬��Ϊ����д��̫����Ӱ��ҳ�����۵ģ�Ҫϸ�µÿ����������������˼��д������Ǳ�Ҫ�ġ�
	<p class="p1">�����ֵ�����������������ѡ�򣬿�������ָ���������еļ�����Ŀ�����Ǹ��ǵ�ԭ�е��ػ�����ϵͳ��ʾ����ȡ���֣���������������Ϊ���޸�������Ŀ��ȷ������ûд��Ļ�������һ��������ѡ��ͬ��ʱ���ʡ�����ȷ���������ݵİ�ȫ��
	<p class="p1">ֻ�й���Ա�ſ��Խ�������������
	
<%case "003"%>
	<p class="p1"><font class="fonts"><b>����޸Ļ�ɾ���Ѿ�����ļ���������</b></font>
	<p class="p1">���޸ģ���ֻ��Ҫ�����µļ����������м������ڼ������ҳ��ĵײ��ı�������ʹ��ԭ�е����ֱ��棬��ѡ��ͬ��ʱ���ǡ����������޸���ԭ���������Ŀ��
	<p class="p1">��ɾ�����ڡ��Զ���ͳ�ơ�ҳ����Զ�����������б��е����ɾ�������ӣ�����ɾ��ҳ�棬ϵͳ��ʾ�Ƿ����Ҫɾ�������ȷ�����ɡ�
	<p class="p1">ֻ�й���Ա�ſ��Խ�������������

<%case "004"%>
	<p class="p1"><font class="fonts"><b>�ÿͺ͹���Ա�ֱ������ЩȨ�ޣ�</b></font>
	<p class="p1">������վ����ͳ�ƶԷ���Ȩ��ʵ�еȼ���������Աʼ�վ��� 6 ����Ȩ������ӵ������Ȩ�����ǹ���Ա���ÿͣ���ӵ�е�Ȩ���ɹ���Ա�ڰ�װʱ�������ļ������á��ȼ����ֵİ취�ǣ�</p>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="320" align=center>
	<tr height="18" align=center class=backzq><td width="110">�� ��</td><td width="30">0</td><td width="30">1</td><td width="30">2</td><td width="30">3</td><td width="30">4</td><td width="30">5</td><td width="30">6</td></tr>
	<tr height="18" align=center><td align=left>&nbsp;�鿴��������</td><td></td><td>��</td><td>��</td><td>��</td><td>��</td><td>��</td><td>��</td></tr> 
	<tr height="18" align=center class=backzq><td align=left>&nbsp;�鿴ȱʡͳ������</td><td></td><td></td><td>��</td><td>��</td><td>��</td><td>��</td><td>��</td></tr> 
	<tr height="18" align=center><td align=left>&nbsp;�鿴��ϸ��¼</td><td></td></td><td></td><td><td>��</td><td>��</td><td>��</td><td>��</td></tr> 
	<tr height="18" align=center class=backzq><td align=left>&nbsp;�鿴�Զ����¼</td><td></td><td></td><td></td><td></td><td>��</td><td>��</td><td>��</td></tr> 
	<tr height="18" align=center><td align=left>&nbsp;�Զ�������ͳ��</td><td></td><td></td><td></td><td></td><td></td><td>��</td><td>��</td></tr> 
	<tr height="18" align=center class=backzq><td align=left>&nbsp;�����Զ����</td><td></td><td></td><td></td><td></td><td></td><td></td><td>��</td></tr> 
</table>
	<p class="p1">����ÿ͵ȼ�������Ϊ 0����ÿͽ�û���κ�Ȩ��������ÿ͵ȼ�������Ϊ 6����ÿͽ����к͹���Ա��ȵ�Ȩ����Ĭ�ϵķÿ�Ȩ�޵ȼ�Ϊ 4��
	<p class="p1">��ǰ�ÿ͵�Ȩ�޵ȼ�Ϊ��<%=whatcan%>

<%case "005"%>
	<p class="p1"><font class="fonts"><b>��λ�ñ�ϵͳ�����°汾��</b></font>
	<p class="p1">������ע�ҵ���ҳ <a href="http://www.ajiang.net" target="_blank">http://www.ajiang.net</a> �һ����Ƴ��°�ĵ�һʱ������ҳ�Ϸ�������Ҳ���Ӱ�������ҳ�ϻ�ð�����д����������
	<p class="p1">�����ע�Ᵽ�������İ�Ȩ��Ϣ���һ��Ѻܶྫ���������ֳ����������ʹ�ã�����һ�ֳɾ͸У��Ҳ��������ȡ������ҫ��
	<p class="p1">����л��ʹ�ð�����д�ĳ���ϣ�����Եõ����Ľ���ͼ���֧�֡�

<%case "006"%>
	<p class="p1"><font class="fonts"><b>��ϵͳ�İ�Ȩ��Ϣ</b></font>
	<p class="p1">������Ϊ��������(shareware)�ṩ������վ���ʹ�ã����߰���(ajiang.net)����ȫ��Ȩ��������ط��ɺ͹��ʹ�Լ����������Ƿ��޸ġ�ת�ء�ɢ��������������Ӯ����Ϊ��������ɾ����Ȩ������
<p class="p1">ϣ���������ñ�ϵͳ���ҵ���վ���ı�����ĸ���֪ͨ�ʼ�����ϵͳ��Լÿ�ܸ���һ�Ρ�
<p class="p1">�������ʹ�������������⣬�뵽�ҵ���վ���ʡ�
<%case "007"%>
	<p class="p1"><font class="fonts"><b>Ϊʲôϵͳ��ʾ�ҡ���ϵͳ��δ���á���</b></font>
	<p class="p1">������Ϊ��ͳ�ƽ������0�����������Ϊ����û���ڱ�ͳ��ҳ�����ͳ�ƴ��룬���߷������Ժ�ͳ��ҳ�滹û�б����ʹ���ͳ��ϵͳ��Ƕ������ǣ�
	<p class="p1">&lt;script src="<%="http://" & Request.ServerVariables("http_host") & finddir(Request.ServerVariables("SCRIPT_NAME"))%>mystat.asp?style=icon"&gt;&lt;/script&gt;
	<p class="p1">�й�Ƕ������д����ο���ϵͳ�İ�װ˵����
	<p class="p1">�����ȷ�����Ѿ���ȷ���������ϴ��벢���ʹ������������ķ������Ƿ�����д�롣
	
<%case else%>
<p class="p1">Ŀǰ��û�������صİ�����Ϣ��

<%end select%>
<p class="p1" align="right"><a href='javascript:history.back()'>����</a> <a href='javascript:history.back()'><img src="images/arbutton.gif" align="absmiddle" border="0"></a> <font style="font-size:16px">&nbsp;</font>&nbsp;
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