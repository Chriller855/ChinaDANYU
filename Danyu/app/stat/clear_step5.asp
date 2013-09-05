<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!-- #include file="inc_config.asp" -->
<%
'权限检查

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
<title><%=countname%>－数据清理－第五步－整理文件</title>
<link rel="stylesheet" type="text/css" href="../../index.css">
</head>
<body topmargin=5 rightmargin=0 leftmargin=0 vlink=#000000>
<table width="80%" cellspacing="0" align="center" cellpadding="0" border="0">
  </tr>
  <tr height="30">
    <td width="1" class="backs"></td>
    <td class="backq">
    
		数据清理－第五步&nbsp; □ 整理文件 ∷∷∷<br>

<table width="90%" align=center><tr>
          <td>
<%if coperr or delerr then%>
<p class="p1">在复制和删除文件时出现错误，请进入FTP查对您的文件，count.mdb和temp.mdb中较小的一个应该保留，较大的删除，保留的一个最终要改名为count.mdb。
<p class="p1" align="right"><a href='#' onclick="window.close()">关闭窗口</a> &nbsp;
  <%else%>
<p class="p1">已经完成了文件的最后整理。
<p class="p1" align="right"><a href='#' onclick="window.close()">关闭窗口</a> &nbsp;
  <%end if%>
</td></tr></table>

	</td>
  </tr>
</table>
</body>
</html>
