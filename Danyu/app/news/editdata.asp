<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->

<!-- #include file="../user/chkpas.asp" -->
<!-- #include file="conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title></title>
<link href="../index.css" rel="stylesheet" type="text/css">
</head>
<%
sortid=clng(request("sortid"))
if sortid=0 then sortid=1
Set Rs=Server.CreateObject("Adodb.Recordset")
Sql="Select top 1 * from newssort where sortid="&sortid
Rs.open Sql,Conn,1,1
sortname=rs("sortname")
attributeid=rs("attributeid")
session("refererString")="editdata.asp?sortid="&sortid
rs.close
set rs=nothing
%>
<body>

<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0" bordercolor="#111111" style="border-collapse: collapse">
  <tr>
    <td width="100%">
<p align="center">新闻管理系统</p>
    </td>
  </tr>
  <tr>
    <td width="100%">
<%if clng(request("sortid"))<>0 and clng(request("sortid"))<>1 then%>
    <form method="POST" action="sortdata.asp" target="_self">
      <p align="left">                <br>
      ·修改栏目</p>
            <p align="center"> 栏目 
              <input type="text" name="sortname" size="30" value="<%=sortname%>">
              &nbsp;所属 
              <!-- #include file="listTitlesAttribute.asp" -->
              <select name="updateTime">
                <option selected>不更新日期</option>
                <option value="1">更新日期</option>
              </select>
              <input type="submit" value="修改" name="B1">
            </p>
      <input type="hidden" name="sortid" value="<%=sortid%>">
    </form><%end if%>
    <p>　</td>
  </tr>
  <tr>
    <td width="100%" height="30"> ·栏目顺序<br>
      <br>
    <table border="1" cellspacing="1" width="100%" id="AutoNumber2" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF">
      <tr>
        <td width="120" align="center">栏目名</td>
        <td align="center">上级栏目</td>
          <td width="200" align="center">下级栏目/顺序</td>
        <!--<td width="10%" align="center">删除</td>-->
      </tr>
      <tr>
        <td><%=sortname%>　 
</td>
        <td>
<%
str=""
while attributeid<>0
Set Rs=Server.CreateObject("Adodb.Recordset")
Sql="Select sortid,sortname,attributeid from newssort where sortid="&attributeid

Rs.open Sql,Conn,1,1
attributeid=rs("attributeid")
str="<a href='?sortid="&rs("sortid")&"'>"&rs("sortname")&"</a> → "&str
rs.close
set rs=nothing
wend
response.write str
%>　</td>
        <td>
<%Set Rs=Server.CreateObject("Adodb.Recordset")
Sql="Select sortid,sortname,sortPos from newssort where attributeid ="&sortid&" order by sortpos desc,createTime desc"
Rs.open Sql,Conn,1,1
%>
					<table width="100%" border="0" cellspacing="0" cellpadding="0">
<%do while not rs.EOF%>
            <form name="form1" method="post" action="changePos.asp?sortid=<%=rs("sortid")%>">
              <tr> 
                <td><%="·<a href='?sortid="&rs("sortid")&"'>"&rs("sortname")&"</a>"%></td>
                <td width="35"><input name="sortPos" type="text" value="<%=rs("sortPos")%>" size="3" maxlength="3"></td>
                <td width="50"><input type="submit" name="Submit" value="修改"></td>
              </tr>
            </form> 
<%
rs.movenext
loop
rs.close
set rs=nothing
%>
					</table></td>
        <!--<td width="10%"><a href="sortdel.asp?sortid=<%=sortid%>">删除</a></td>-->
      </tr>
    </table>
    </td>
  </tr>
  </table>

<% 
Conn.Close
Set Conn=Nothing%>
</body>
</html>
