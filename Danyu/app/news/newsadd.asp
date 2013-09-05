<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!-- #include file="conn.asp" -->
<html>
<head>
<title>新闻添加</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../index.css" rel="stylesheet" type="text/css">
<script language="JavaScript">
<!--
function _save(){
	if (document.hiddenedform.news_title.value==""){
		alert("新闻标题不能为空！");
		return false;
	}
	// getHTML()为eWebEditor自带的接口函数，功能为取编辑区的内容
	//if (eWebEditor1.getHTML()==""){
	//	alert("新闻内容不能为空！");
	//	return false;
	//}
	hiddenedform.submit();
}
-->
</SCRIPT>
</head>
<body>
<%
if request("action")="add" then
	Set Rs=Server.CreateObject("Adodb.Recordset")
	Sql="Select * from article"
	Rs.open Sql,Conn,3,4
	rs.addnew
	rs("sortid")=request.form("sortid")
	if request("updateTime")="now" then rs("news_uptime")=now()
	if request("updateTime")="refer" then
		rs("news_uptime")=cdate(request("year")&"-"&request("month")&"-"&request("day")&" "&time())
	end if
	rs("news_title")=request("news_title")
	rs("news_SubTitle")=request("news_SubTitle")
	rs("news_level")=1
	rs("news_click")=0
	rs("news_author")=request("Author")
	rs("news_From")=request("comeFrom")
	rs("news_istop")=request("istop")
	rs("news_tpic")=request("PicUrl")
	if trim(split(request("accessRight"),",")(0))="0" then
		rs("accessRight")=null
	else
		rs("accessRight")=request("accessRight")
	end if
	rs("news_Url")=request("newsUrl")
	rs("news_original")=request("original")
	rs("news_editor")=session("loginFlag")
	
	'转换绝对路径
	http1="http://"&Request.ServerVariables("SERVER_NAME")
	http2=Request.ServerVariables("SERVER_PORT")
	if http2<>"80" then
	http1=http1&":"&http2
	end if
	'response.write http1
	
	' 注意取新闻内容的方法，因为对大表单的自动处理，一定要使用循环，否则大于100K的内容将取不到，单个表单项的限制为102399字节（100K左右）
	' 开始：eWebEditor编辑区取值-----------------
	Content = ""
	For i = 1 To Request.Form("news_Content").Count
		Content = Content & Request.Form("news_Content")(i)
	Next
	' 结束：eWebEditor编辑区取值-----------------
	
	Content=replace(Content,http1,"")
	rs("news_content")=Content
	''''''''''''''''''''''''''''''''''''''
	
	rs.updatebatch
	rs.close
	set rs=nothing
	Conn.Close
	Set Conn=Nothing
	response.write "<p style=font-size:12pt; align=center>恭喜您，新闻“<b>"&request("news_title")&"</b>”已经成功修改!<br>"
	response.write "<a href='javascript:window.close()'><font style=font-size:12pt;>关闭</font></a>"
else
%>
<form name="hiddenedform" method="post" action="?action=add">
  <table width="100%" border="0" cellspacing="0" cellpadding="2" bordercolordark="#FFFFFF">
    <tr> 
      <td width="100" align="right">所属栏目：</td>
      <td><!-- #include file="listTitlesSort.asp" --></td>
    </tr>
<tr> 
      <td align="right">是否指定时间：</td>
      <td><script language="JavaScript">
	  function enableInput(v){
	  	hiddenedform.year.disabled=v;
	  	hiddenedform.month.disabled=v;
	  	hiddenedform.day.disabled=v;
	  }
	  </script>
        <input name="updateTime" type="radio" onClick="javascript:enableInput(true)" value="now" checked>当前时间
        <input name="updateTime" type="radio" value="refer"  onClick="javascript:enableInput(false)">
        指定（ 
        <select name="year" style="width:55px;" disabled>
          <option value="<%=year(now())%>" selected><%=year(now())%></option>
<%for i=2005 to 2020%>
          <option value="<%=i%>"><%=i%></option>
<%next%>
        </select>
        年<select name="month" style="width:38px;" disabled>
          <option value="<%=month(now())%>" selected><%=month(now())%></option>
<%for i=1 to 12%>
          <option value="<%=i%>"><%=i%></option>
<%next%>
        </select>
        月
<select name="day" style="width:38px;" disabled>
          <option value="<%=day(now())%>" selected><%=day(now())%></option>
<%for i=1 to 31%>
          <option value="<%=i%>"><%=i%></option>
<%next%>
        </select>
        日）</td>
    </tr>
	    <tr> 
      <td align="right">题目（项目名）：</td>
      <td><input name="news_title" type="text" id="title" size="50"></td>
    </tr>
    <tr> 
      <td align="right">副标题（标题语）：</td>
      <td><input name="news_SubTitle" type="text" id="title" size="50"></td>
    </tr>
    <tr> 
      <td align="right">作者（开发地区）：</td>
      <td> <input name="Author" size=20 maxlength=50 title='如果有，在这里输入文章作者'>
        摘自（楼盘电话）：
<input name="comeFrom" size=21 maxlength=50 title='如果有，在这里输入文章作者'></td>
    </tr>
    <tr> 
      <td align="right"> 标题图片URL：</td>
      <td> 
<input name="PicUrl" size="20" maxlength="100">
        文章固顶： 
        <select size="1" name="istop">
          <option value="0" selected>0</option>
		  <%for i=1 to 10%>
          <option value="<%=i%>"><%=i%></option>
		  <%next%>
        </select>（0表示按时间倒序）<BR><IFRAME name=ad src="../uploadfile/insertpic.asp?obj=PicUrl&picFlag=0&frmName=hiddenedform" frameBorder=0 width="300" scrolling=no height=25></IFRAME>
    </tr>
    <tr> 
      <td align="right"> 新闻转向URL：</td>
      <td> <input name="newsUrl" type="text" size="40" maxlength="100">
        <br>
        <IFRAME name=ad src="../uploadfile/insertpic.asp?obj=newsUrl&picFlag=0&frmName=hiddenedform" frameBorder=0 width="300" scrolling=no height=25></IFRAME>
    </tr>
    <tr> 
      <td align="right">简介（楼盘地址）<br>
        (200字内)：</td>
      <td><table width="92%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="250"><textarea rows="7" name="original" cols="30"></textarea></td>
            <td>文章权限：<br>
              (按住CTRL键点击复选)<br> 
              <%
response.write "<select name='accessRight' size='5' style='width:110px;' multiple>"
Set Rs2=Server.CreateObject("Adodb.Recordset")
Sql="Select * from userTypeInfo order by typeID asc"
Rs2.open Sql,Conn,1,1
response.write "<option value='0' selected>所有浏览者</option>"
do while not rs2.eof
	response.write "<option value='"&rs2("typeID")&"'>"
	response.write rs2("typeName")
	response.write "</option>"
	rs2.movenext
loop
response.write "</select>"
rs2.close
set rs2=nothing
%>
            </td>
          </tr>
        </table> </td>
    </tr>
    <tr> 
      <td align="right">内容：</td>
      <td> 
		<input type=hidden name=news_Content>
		<IFRAME ID="eWebEditor1" src="eWebEditor/ewebeditor.asp?id=news_Content&style=YOUNG" frameborder="0" scrolling="no" width="600" height="350"></IFRAME>
      </td>
    </tr>
    <tr> 
      <td align="right">&nbsp;</td>
      <td><input name="Input" type="button" value="    提  交    " onclick='_save()'></td>
    </tr>
  </table>
</form>
<%
end if
%>
</body>
</html>