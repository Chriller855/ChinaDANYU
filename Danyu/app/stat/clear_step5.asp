<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!-- #include file="inc_config.asp" -->
<%
'Ȩ�޼��

dbPath=server.MapPath(connpath)
strDBPath = left(DBPath,instrrev(DBPath,"\")) 

Set fso = CreateObject("Scripting.FileSystemObject") 
on error resume next
fso.CopyFile strDBPath & "temp.mdb",dbpath 
if err<>0 then coperr=true
fso.DeleteFile(strDBPath & "temp.mdb") 
if err<>0 then delerr=true
Set fso = nothing 

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<title><%=countname%>�������������岽�������ļ�</title>
<link rel="stylesheet" type="text/css" href="../../index.css">
</head>
<body topmargin=5 rightmargin=0 leftmargin=0 vlink=#000000>
<table width="80%" cellspacing="0" align="center" cellpadding="0" border="0">
  </tr>
  <tr height="30">
    <td width="1" class="backs"></td>
    <td class="backq">
    
		�����������岽&nbsp; �� �����ļ� �ˡˡ�<br>

<table width="90%" align=center><tr>
          <td>
<%if coperr or delerr then%>
<p class="p1">�ڸ��ƺ�ɾ���ļ�ʱ���ִ��������FTP��������ļ���count.mdb��temp.mdb�н�С��һ��Ӧ�ñ������ϴ��ɾ����������һ������Ҫ����Ϊcount.mdb��
<p class="p1" align="right"><a href='#' onclick="window.close()">�رմ���</a> &nbsp;
  <%else%>
<p class="p1">�Ѿ�������ļ����������
<p class="p1" align="right"><a href='#' onclick="window.close()">�رմ���</a> &nbsp;
  <%end if%>
</td></tr></table>

	</td>
  </tr>
</table>
</body>
</html>
