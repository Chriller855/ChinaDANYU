<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!-- #include file="conn.asp" -->
<html>
<head>
<title>�����޸�</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../index.css" rel="stylesheet" type="text/css">
<script language="JavaScript">
<!--
function _save(){
	if (document.hiddenedform.news_title.value==""){
		alert("���ű��ⲻ��Ϊ�գ�");
		return false;
	}
	// getHTML()ΪeWebEditor�Դ��Ľӿں���������Ϊȡ�༭��������
	//if (eWebEditor1.getHTML()==""){
	//	alert("�������ݲ���Ϊ�գ�");
	//	return false;
	//}
	hiddenedform.submit();
}
-->
</SCRIPT>
</head>
<body>
<%
if request("action")="modify" then
	Set Rs=Server.CreateObject("Adodb.Recordset")
	Sql="Select * from article where id="&request("id")
	Rs.open Sql,Conn,3,4
	rs("sortid")=request.form("sortid")
	if request("updateTime")="now" then rs("news_uptime")=now()
	if request("updateTime")="refer" then
		rs("news_uptime")=cdate(request("year")&"-"&request("month")&"-"&request("day")&" "&time())
	end if
	rs("news_title")=request("news_title")
	rs("news_SubTitle")=request("news_SubTitle")
	rs("news_author")=request("Author")
	rs("news_from")=request("comeFrom")
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
	
	'ת������·��
	http1="http://"&Request.ServerVariables("SERVER_NAME")
	http2=Request.ServerVariables("SERVER_PORT")
	if http2<>"80" then
	http1=http1&":"&http2
	end if
	'response.write http1
	
	' ע��ȡ�������ݵķ�������Ϊ�Դ�����Զ�����һ��Ҫʹ��ѭ�����������100K�����ݽ�ȡ�������������������Ϊ102399�ֽڣ�100K���ң�
	' ��ʼ��eWebEditor�༭��ȡֵ-----------------
	Content = ""
	For i = 1 To Request.Form("news_Content").Count
		Content = Content & Request.Form("news_Content")(i)
	Next
	' ������eWebEditor�༭��ȡֵ-----------------

	Content=replace(Content,http1,"")
	rs("news_content")=Content
	''''''''''''''''''''''''''''''''''''''
	
	rs.updatebatch
	rs.close
	set rs=nothing
	Conn.Close
	Set Conn=Nothing
	response.write "<p style=font-size:12pt; align=center>��ϲ�������š�<b>"&request("news_title")&"</b>���Ѿ��ɹ��޸�!<br>"
	response.write "<a href='javascript:window.close()'><font style=font-size:12pt;>�ر�</font></a>"
else
	Set Rs=Server.CreateObject("Adodb.Recordset")
	Sql="Select * from article where id="&request("id")
	Rs.open Sql,Conn,1,1
%>
<form name="hiddenedform" method="post" action="?action=modify&id=<%=request("id")%>">
<input type=hidden name=pagebody>
  <table width="100%" border="0" cellspacing="0" cellpadding="3">
    <tr> 
      <td width="100" align="right">������Ŀ��</td>
      <td><!-- #include file="listTitlesSort.asp" --></td>
    </tr>
    <tr> 
      <td align="right">�Ƿ����ʱ�䣺</td>
      <td><script language="JavaScript">
	  function enableInput(v){
	  	hiddenedform.year.disabled=v;
	  	hiddenedform.month.disabled=v;
	  	hiddenedform.day.disabled=v;
	  }
	  </script> <input name="updateTime" type="radio" value="no"  onClick="javascript:enableInput(true)" checked>
        �� 
        <input name="updateTime" type="radio" value="now" onClick="javascript:enableInput(true)">
        ��ǰʱ�� 
        <input name="updateTime" type="radio" value="refer"  onClick="javascript:enableInput(false)">
        ָ���� 
        <select name="year" style="width:55px;" disabled>
          <option value="<%=year(rs("news_uptime"))%>" selected><%=year(rs("news_uptime"))%></option>
          <%for i=2003 to 2010%>
          <option value="<%=i%>"><%=i%></option>
          <%next%>
        </select>
        �� 
        <select name="month" style="width:38px;" disabled>
          <option value="<%=month(rs("news_uptime"))%>" selected><%=month(rs("news_uptime"))%></option>
          <%for i=1 to 12%>
          <option value="<%=i%>"><%=i%></option>
          <%next%>
        </select>
        �� 
        <select name="day" style="width:38px;" disabled>
          <option value="<%=day(rs("news_uptime"))%>" selected><%=day(rs("news_uptime"))%></option>
          <%for i=1 to 31%>
          <option value="<%=i%>"><%=i%></option>
          <%next%>
        </select>
        �գ�</td>
    </tr>
    <tr> 
      <td align="right">��Ŀ����Ŀ������</td>
      <td><input name="news_title" type="text" id="title" value="<%=rs("news_title")%>" size="50"></td>
    </tr>
    <tr> 
      <td align="right">�����⣨�������</td>
      <td><input name="news_SubTitle" type="text" id="title" value="<%=rs("news_SubTitle")%>" size="50"></td>
    </tr>
    <tr> 
      <td align="right">���ߣ�������������</td>
      <td> <input value="<%=rs("news_author")%>" name="Author" size=20 maxlength=50 title='����У�������������������'>
        ժ�ԣ�¥�̵绰���� 
        <input value="<%=rs("news_from")%>" name="comeFrom" size=21 maxlength=50 title='����У�������������������'></td>
    </tr>
    <tr> 
      <td align="right"> ����ͼƬURL��</td>
      <td> <input name="PicUrl" type="text" size="20" maxlength="100" value="<%=rs("news_tpic")%>">
        ���¹̶��� 
        <select size="1" name="istop">
          <%for i=0 to 10%>
          <option value="<%=i%>" <%if rs("news_istop")=i then%>selected<%end if%>><%=i%></option>
          <%next%>
        </select>
        ��0��ʾ��ʱ�䵹�� <BR> <IFRAME name=ad src="../uploadfile/insertpic.asp?obj=PicUrl&picFlag=0&frmName=hiddenedform" frameBorder=0 width="300" scrolling=no height=25></IFRAME> 
      </td>
    </tr>
    <tr> 
      <td align="right"> ����ת��URL��</td>
      <td> <input name="newsUrl" type="text" size="50" maxlength="100" value="<%=rs("news_Url")%>"> 
        <br> <IFRAME name=ad src="../uploadfile/insertpic.asp?obj=newsUrl&picFlag=0&frmName=hiddenedform" frameBorder=0 width="300" scrolling=no height=25></IFRAME> 
    </tr>
    <tr> 
      <td align="right">��飨¥�̵�ַ��<br>
        (200����)��</td>
      <td> <table width="92%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="250"><textarea rows="7" name="original" cols="30"><%=rs("news_original")%></textarea></td>
            <td>����Ȩ�ޣ�<br>
              (��סCTRL�������ѡ)<br> <%
response.write "<select name='accessRight' size='5' style='width:110px;' multiple>"
Set Rs2=Server.CreateObject("Adodb.Recordset")
Sql="Select * from userTypeInfo order by typeID asc"
Rs2.open Sql,Conn,1,1
response.write "<option value='0'"
if isNull(rs("accessRight")) then
	response.write " selected"
	temp=split("",",")
else
	if rs("accessRight")=""  then response.write " selected"
	temp=split(rs("accessRight"),",")
end if
response.write ">���������</option>"
do while not rs2.eof
	response.write "<option value='"&rs2("typeID")&"'"
	for i=0 to ubound(temp)
		if trim(rs2("typeID"))=trim(temp(i)) then
			response.write " selected"
			exit for
		end if
	next
	response.write ">"
	response.write rs2("typeName")
	response.write "</option>"
	rs2.movenext
loop
response.write "</select>"
rs2.close
set rs2=nothing
%> </td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td align="right">���ݣ�</td>
      <td><input name="news_Content" type="hidden" value="<%=Server.HtmlEncode(rs("news_Content"))%>"> 
        <IFRAME ID="eWebEditor1" src="eWebEditor/eWebEditor.asp?id=news_Content&style=YOUNG" frameborder="0" scrolling="no" width="600" height="350"></IFRAME> 
      </td>
    </tr>
    <tr> 
      <td align="right">&nbsp;</td>
      <td height="32"><input name="Input" type="button" value="    ��  ��    " onClick='_save()'></td>
    </tr>
  </table>
</form>
<%
Conn.Close
Set Conn=Nothing
end if
%>
</body>
</html>