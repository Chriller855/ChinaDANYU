<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!-- #include file="conn.asp" -->
<html>
<head>
<title>�������</title>
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
%>
<form name="hiddenedform" method="post" action="?action=add">
  <table width="100%" border="0" cellspacing="0" cellpadding="2" bordercolordark="#FFFFFF">
    <tr> 
      <td width="100" align="right">������Ŀ��</td>
      <td><!-- #include file="listTitlesSort.asp" --></td>
    </tr>
<tr> 
      <td align="right">�Ƿ�ָ��ʱ�䣺</td>
      <td><script language="JavaScript">
	  function enableInput(v){
	  	hiddenedform.year.disabled=v;
	  	hiddenedform.month.disabled=v;
	  	hiddenedform.day.disabled=v;
	  }
	  </script>
        <input name="updateTime" type="radio" onClick="javascript:enableInput(true)" value="now" checked>��ǰʱ��
        <input name="updateTime" type="radio" value="refer"  onClick="javascript:enableInput(false)">
        ָ���� 
        <select name="year" style="width:55px;" disabled>
          <option value="<%=year(now())%>" selected><%=year(now())%></option>
<%for i=2005 to 2020%>
          <option value="<%=i%>"><%=i%></option>
<%next%>
        </select>
        ��<select name="month" style="width:38px;" disabled>
          <option value="<%=month(now())%>" selected><%=month(now())%></option>
<%for i=1 to 12%>
          <option value="<%=i%>"><%=i%></option>
<%next%>
        </select>
        ��
<select name="day" style="width:38px;" disabled>
          <option value="<%=day(now())%>" selected><%=day(now())%></option>
<%for i=1 to 31%>
          <option value="<%=i%>"><%=i%></option>
<%next%>
        </select>
        �գ�</td>
    </tr>
	    <tr> 
      <td align="right">��Ŀ����Ŀ������</td>
      <td><input name="news_title" type="text" id="title" size="50"></td>
    </tr>
    <tr> 
      <td align="right">�����⣨�������</td>
      <td><input name="news_SubTitle" type="text" id="title" size="50"></td>
    </tr>
    <tr> 
      <td align="right">���ߣ�������������</td>
      <td> <input name="Author" size=20 maxlength=50 title='����У�������������������'>
        ժ�ԣ�¥�̵绰����
<input name="comeFrom" size=21 maxlength=50 title='����У�������������������'></td>
    </tr>
    <tr> 
      <td align="right"> ����ͼƬURL��</td>
      <td> 
<input name="PicUrl" size="20" maxlength="100">
        ���¹̶��� 
        <select size="1" name="istop">
          <option value="0" selected>0</option>
		  <%for i=1 to 10%>
          <option value="<%=i%>"><%=i%></option>
		  <%next%>
        </select>��0��ʾ��ʱ�䵹��<BR><IFRAME name=ad src="../uploadfile/insertpic.asp?obj=PicUrl&picFlag=0&frmName=hiddenedform" frameBorder=0 width="300" scrolling=no height=25></IFRAME>
    </tr>
    <tr> 
      <td align="right"> ����ת��URL��</td>
      <td> <input name="newsUrl" type="text" size="40" maxlength="100">
        <br>
        <IFRAME name=ad src="../uploadfile/insertpic.asp?obj=newsUrl&picFlag=0&frmName=hiddenedform" frameBorder=0 width="300" scrolling=no height=25></IFRAME>
    </tr>
    <tr> 
      <td align="right">��飨¥�̵�ַ��<br>
        (200����)��</td>
      <td><table width="92%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="250"><textarea rows="7" name="original" cols="30"></textarea></td>
            <td>����Ȩ�ޣ�<br>
              (��סCTRL�������ѡ)<br> 
              <%
response.write "<select name='accessRight' size='5' style='width:110px;' multiple>"
Set Rs2=Server.CreateObject("Adodb.Recordset")
Sql="Select * from userTypeInfo order by typeID asc"
Rs2.open Sql,Conn,1,1
response.write "<option value='0' selected>���������</option>"
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
      <td align="right">���ݣ�</td>
      <td> 
		<input type=hidden name=news_Content>
		<IFRAME ID="eWebEditor1" src="eWebEditor/ewebeditor.asp?id=news_Content&style=YOUNG" frameborder="0" scrolling="no" width="600" height="350"></IFRAME>
      </td>
    </tr>
    <tr> 
      <td align="right">&nbsp;</td>
      <td><input name="Input" type="button" value="    ��  ��    " onclick='_save()'></td>
    </tr>
  </table>
</form>
<%
end if
%>
</body>
</html>