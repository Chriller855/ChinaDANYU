<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!-- #include file="inc_config.asp" -->
<%
wherestr=Request("wherestr")

pagesize=mPageSize
if request("page")="" then
  	curpage = 1
else
	curpage = cint(request("page"))
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<title><%=countname%>－详细记录</title>
<link rel="stylesheet" type="text/css" href="../../index.css">
</head>
<body topmargin=5 rightmargin=0 leftmargin=0 vlink=#000000>
详细记录<br>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="98%" align=center>
  <tr align=center height="25">
    <td width="85">时 &nbsp; 间</td>
    <td width="75">地区</td>
    <td width="40">屏宽</td>
    <td width="55">操作系统</td>
    <td width="55">浏览器</td>
    <td width="170">来 源 网 页</td>
  </tr>
  <%
Set rs = Server.CreateObject("ADODB.Recordset")

sql = "SELECT * FROM view " & wherestr & " ORDER BY id DESC"
rs.open sql, conn, 1, 1
if rs.bof and rs.eof then
	rs.close
	response.write "<tr><td colspan='6'><br><center>" & wherestr & "没有符合条件的记录</center><br></td></tr></table></td></tr></table>"
else
	dim i
	rs.pagesize = pagesize
	if rs.pagecount < curpage then
		rs.absolutepage=rs.pagecount
		curpage=rs.pagecount
	else
		rs.absolutepage = curpage
	end if
for i = 1 to rs.pagesize
%>
  <tr height="18">
    <td><a title="<%=rs("vpage")%>"><%=month(rs("vtime")) & "-" & day (rs("vtime")) & " " & formatdatetime(rs("vtime"),4)%></a></td>
    <td><font face="Times new roman"><a title="<%=rs("vIP") & " " & rs("vwhere") & "·" & rs("vwheref")%>"><%=rs("vwhere")%></a></font></td>
    <td align=center><font face="Times new roman"><%=rs("vwidth")%></font></td>
    <td align=center><font face="Times new roman"><%=rs("vOS")%></font></td>
    <td align=center><font face="Times new roman"><%=rs("vsoft")%></font></td>
    <td><%
	vcome=rs("vcome")
	thelen=len(vcome)
	if thelen=0 then Response.Write ""
	if thelen <= 33 and thelen > 0 then
		svcome=right(vcome,thelen-6)
		Response.Write "<a href='" & vcome & "' target='_blank'>" & svcome & "</a>"
	end if
	if thelen >= 34 then
		svcome=left(right(vcome,thelen-6),24) & "..."
		Response.Write "<a title='" & vcome & "' href='" & vcome & "' target='_blank'>" & svcome & "</a>"
	end if
	%></td>
  </tr>
  <%
  rs.movenext
    if rs.eof then
      i = i + 1
      exit for
    end if
  next
%>
</table>
<br>
用鼠标点指地区可看到IP地址及详细地区，点指访问日期可查看被访问页面。<br>
<br>
<table width=98% align=center cellpadding=0 cellspacing=0>
  <form method=post action="tj_all.asp" id=form2 name=form2>
    <tr height=65>
      <td width="498"> 第<font class="fonts"><%=cstr(curpage)%></font>页 总<font class="fonts"><%=cstr(rs.pagecount)%></font>页 本页<font class="fonts"><%=cstr(i-1)%></font>条 总<font class="fonts"><%=cstr(rs.recordcount)%></font>条 &nbsp;
        <%
      wherestr=server.URLEncode(wherestr)
      if curpage <> 1 then%>
        <a href="tj_all.asp?page=1&wherestr=<%=wherestr%>">&lt;&lt;首页 </a> <a href="tj_all.asp?page=<%=cstr(curpage-1)%>&wherestr=<%=wherestr%>">&lt;上一页 </a>
        <%else%>
        <font color="#888888">&lt;&lt;首页 &lt;上一页</font>
        <%end if
      if curpage <> rs.pagecount then%>
        <a href="tj_all.asp?page=<%=cstr(curpage+1)%>&wherestr=<%=wherestr%>"> 下一页&gt;</a> <a href="tj_all.asp?page=<%=cstr(rs.pagecount)%>&wherestr=<%=wherestr%>"> 尾页&gt;&gt;</a>
        <%else%>
        <font color="#888888">下一页&gt; 尾页&gt;&gt;</font>
        <%end if%>
        <input type=hidden name=wherestr value="<%=wherestr%>">
&nbsp;
        <input type=text name='page' size=9 class=input>
        <input type=submit value='GO'>
      </td>
    </tr>
  </form>
</table>
<br>
<%
'如果是检索结果，则显示保存数据表单
if wherestr <> "" then
wherestr=Request("wherestr")
'保存数据表单
if session.Contents("master")=true or whatcan>=6 then

%>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
  <form action="tj_save.asp" method=post id=form1 name=form1>
    <tr height="30">
      <td width="498"> 保存这次检索条件 ∷∷∷<br>
          <p class="p1"><font class="fonts"><b>检</b></font><b>索条件</b>&nbsp;
              <%if wherestr="" then%>
              没有检索条件
              <%else%>
              <%=wherestr%>
              <%end if%>
              。
        <p style='margin-bottom: 0;margin-top: 0;'><b>查询项目</b>&nbsp; 详细。
        <p class="p1" style='line-height: 100%;margin-bottom: 0;margin-top: 0;'><font class="fonts"><b>取</b></font><b>个名字</b>&nbsp;
                <input name="name" size="16" class="input">
      &nbsp;
                <INPUT type="radio" name="overwrite" value="0" checked>
          同名时提示&nbsp;
          <INPUT type="radio" name="overwrite" value="1">
          同名时覆盖
        <b>加个介绍</b>&nbsp;
                <input name="content" size="50" class="input">
            <a href="help.asp?id=002">帮助</a> <a href='javascript:document.forms[0].submit();'>保存</a>
                <input type="hidden" name="wherestr" value="<%=wherestr%>">
                <input type="hidden" name="outtype" value="详细">
                <input type="submit" value=" " name="save" class="backc2">
      </td>
    </tr>
  </form>
</table>
<br>
<%
end if
end if
%>

<%
rs.close
end if

set rs=nothing
conn.Close 
set conn=nothing
%>
</body>
</html>
