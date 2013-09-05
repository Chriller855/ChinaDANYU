<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!-- #include file="inc_config.asp" -->
<%
'权限检查

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
<title><%=countname%>－数据清理－第四步－压缩</title>
<link rel="stylesheet" type="text/css" href="../../index.css">
</head>
<body topmargin=5 rightmargin=0 leftmargin=0 vlink=#000000>
<table width="80%" cellspacing="0" align="center" cellpadding="0" border="0">
  <tr height="30">
    <td width="1" class="backs"></td>
    <td class="backq">
    
		数据清理－第四步&nbsp; □ 压缩数据库 ∷∷∷<br>

<table width="90%" align=center><tr>
          <td>
<p class="p1">已经完成了数据库的压缩，现在有两个库文件，<%=DBPath%>是经过清理但没有压缩过的库文件，<%=strdbpath & "temp.mdb"%>是清理后经过压缩的库文件。下一步的操作是将压缩过的库文件复制为<%=DBPath%>，然后删除temp.mdb。
<%if IsObjInstalled("Scripting.FileSystemObject") then%>
<p class="p1">请点击下一步完成该操作。
<p class="p1" align="right"><a href='clear_step5.asp'>下一步 整理文件 开始</a>&nbsp;
<%else%>
<p class="p1">因为您的主机不支持FSO组件，所以这一步必须您手工完成。请您登录FTP，找到<%=DBPath%>这个文件，删除它（最好在删除前先下载它），然后将<%=strdbpath & "temp.mdb"%>改名为<%=DBPath%>。
<p class="p1" align="right"><a href='#' onclick="window.close()">关闭窗口</a>&nbsp;
<%end if%>
</td></tr></table>

	</td>
  </tr>
</table>
</body>
</html>
<%
'组件检查函数
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
