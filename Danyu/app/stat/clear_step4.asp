<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!-- #include file="inc_config.asp" -->
<%
'Ȩ�޼��

Set Engine = CreateObject("JRO.JetEngine") 

dbPath=server.MapPath(connpath)
strDBPath = left(DBPath,instrrev(DBPath,"\")) 
Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBPath, _ 
"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb" 
Set Engine = nothing 

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<title><%=countname%>�������������Ĳ���ѹ��</title>
<link rel="stylesheet" type="text/css" href="../../index.css">
</head>
<body topmargin=5 rightmargin=0 leftmargin=0 vlink=#000000>
<table width="80%" cellspacing="0" align="center" cellpadding="0" border="0">
  <tr height="30">
    <td width="1" class="backs"></td>
    <td class="backq">
    
		�����������Ĳ�&nbsp; �� ѹ�����ݿ� �ˡˡ�<br>

<table width="90%" align=center><tr>
          <td>
<p class="p1">�Ѿ���������ݿ��ѹ�����������������ļ���<%=DBPath%>�Ǿ�������û��ѹ�����Ŀ��ļ���<%=strdbpath & "temp.mdb"%>������󾭹�ѹ���Ŀ��ļ�����һ���Ĳ����ǽ�ѹ�����Ŀ��ļ�����Ϊ<%=DBPath%>��Ȼ��ɾ��temp.mdb��
<%if IsObjInstalled("Scripting.FileSystemObject") then%>
<p class="p1">������һ����ɸò�����
<p class="p1" align="right"><a href='clear_step5.asp'>��һ�� �����ļ� ��ʼ</a>&nbsp;
<%else%>
<p class="p1">��Ϊ����������֧��FSO�����������һ���������ֹ���ɡ�������¼FTP���ҵ�<%=DBPath%>����ļ���ɾ�����������ɾ��ǰ������������Ȼ��<%=strdbpath & "temp.mdb"%>����Ϊ<%=DBPath%>��
<p class="p1" align="right"><a href='#' onclick="window.close()">�رմ���</a>&nbsp;
<%end if%>
</td></tr></table>

	</td>
  </tr>
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
