<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!-- #include file="inc_config.asp" -->
<%
'Ȩ�޼��
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<title><%=countname%>������������һ����׼��</title>
<link rel="stylesheet" type="text/css" href="../../index.css">
</head>
<body topmargin=5 rightmargin=0 leftmargin=0 vlink=#000000>
<table width="80%" cellspacing="0" align="center" cellpadding="0" border="0">
<form action="clear_step2.asp">
  <tr height="30">
    <td width="1" class="backs"></td>
    <td>
    
		����������һ��&nbsp; �� ׼�� �ˡˡ�<br>

<table width="90%" align=center><tr><td>
<p class="p1">�������ݵĲ����������ģ�����ϵͳ�������Ҫ�������������Ϣ������Ӧ�ļ�¼��Ȼ����ʾ��ʹ�����ݿ�ѹ�����߶Կ��ļ�����ѹ����
<p class="p1">1���ڶ����ݿ��������֮ǰ������ȷ�����Ѿ�����Ҫ��������ݱ��ݣ��������Ѿ�ȷ���㲻��Ҫ����ͳ�Ʊ�����Զ���ͳ����ʹ����Щ��Ϣ�ˣ���Ϊ�����������޷��ָ��ġ�
<p class="p1">2��������ķ�������û�а�װFSO�����������޷���ѹ���������ݿ⸲��ԭ�������ݿ⣬Ҳ�޷�ɾ����ʱ���ļ�����ͱ���׼����FTP����ȿ��Կ�������ϵͳ�������취����������ɺ�����ļ��ĸ�������<%
if IsObjInstalled("Scripting.FileSystemObject") then
%>���Ŀռ�֧��FSO��<%
else%><font color=red>��Ŀռ䲻֧��FSO��<%end if%>��
<p class="p1"><font color=red>ע�⣺</font>����һ���������ʼ�¼���������ȥ�����޷����κ�ͳ�Ʊ����п�����Щ��¼������˵���������ڡ���ϸ��¼���п����������������е�ͳ�Ʊ����ж����������Զ���ͳ�ƹ���Ҳ�޷���ѯ����������Ĳ��֡�ʹ�á��󱸿�鿴�������Բ鿴�����ݹ������ݣ����ǹ��ܺ����ޣ�����Ӧ���ڱ�֤�����������е�����£��������������ϸ��¼����Ŀǰ�󱸿�鿴����δ��ɣ�</p>
<p class="p1">׼�����Ժ�������Ҫ��������ݵ������������ʼ��
<%if Request("offtime")<>"" then%>
<p class="p1">���� <input name="offyear" class="input" size="5" value="<%=year(Request("offtime"))%>"> ��
		<input name="offmonth" class="input" size="3" value="<%=month(Request("offtime"))%>"> ��
		<input name="offday" class="input" size="3" value="<%=day(Request("offtime"))%>"> ��ǰ�����ݡ�<INPUT type="checkbox" name=force value=1>�Ƿ�ǿ������
<%else%>
<p class="p1">���� <input name="offyear" class="input" size="5"> ��
		<input name="offmonth" class="input" size="3"> ��
		<input name="offday" class="input" size="3"> ��ǰ�����ݡ�<INPUT type="checkbox" name=force value=1>�Ƿ�ǿ������
<%end if%>
<p class="p1">ע�⣺��ѡ���ˡ�ǿ������ʱ��û�б��ݹ�������Ҳ�������������á������ѡ��������ݿ��Ѿ������޷����ݵ������ʹ�á�
<p class="p1" align="right"><a href='javascript:document.forms[0].submit();'>��һ�� �������� ��ʼ</a> <input type="submit" name="bakgo" value="GO"><font style="font-size:16px">&nbsp;</font>&nbsp;

</td></tr></table>

	</td>
  </tr>
</form>
</table>
</body>
</html>
<%
'�����麯��
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
