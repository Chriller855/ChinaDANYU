<!-- www.00ok.net ��Ȩ���� -->
<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!--#include file="conn.asp"-->
<title>��Ա�б�</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../../index.css" rel="stylesheet" type="text/css">
<head>
<script language=javascript>
function del(id){
	if(confirm('ȷ��Ҫɾ����')){
		document.location.href="del.asp?subid="+id;
	}
}
</script>
</head>
<body leftmargin="0" topmargin="4">
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td> ��ȡ�����б�<br> <%
set rs=server.createobject("adodb.recordset")
sql="select * from visitApp where comeDate='wantBook' order by id desc"
rs.open sql,conn,1,1
if rs.eof then
	response.write "û����ȡ������Ϣ��"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
totalrec=rs.recordcount
'response.write totalrec
const maxPerPage=20
dim n,p,ii
dim currentpage,totalrec,page_count
currentPage=request("page")
if currentpage="" or not isnumeric(currentpage) then
	currentpage=1
else
	currentpage=clng(currentpage)
end if

if totalrec mod maxPerPage=0 then
	n= totalrec \ maxPerPage
else
	n= totalrec \ maxPerPage+1
end if
RS.MoveFirst
if currentpage > n then currentpage = n
if currentpage<1 then currentpage=1
RS.Move (currentpage-1) * maxPerPage
session("visitAppRefererString")=request.ServerVariables("SCRIPT_NAME")&"?page="&Page
%> 
      <table width="100%" border="1" cellspacing="0" cellpadding="2" bordercolordark="#FFFFFF">
        <tr bgcolor="#CCCCCC"> 
          <td>�� ��</td>
          <td>�ʼĵ�ַ</td>
          <td width="50">��������</td>
          <td width="70">��ϵ�绰</td>
          <td width="70">�ֻ�����</td>
          <td> �����ʼ�</td>
          <td>�������</td>
          <td width="35">����</td>
          <td width="35">ɾ��</td>
        </tr>
        <%
do while not rs.eof and page_count<maxPerPage
page_count=page_count+1
%>
        <tr> 
          <td><%=rs("memberName")%>&nbsp;</td>
          <td><%=rs("address")%>&nbsp;</td>
          <td><%=rs("postCode")%>&nbsp;</td>
          <td><%=rs("telphone")%></td>
          <td><%=rs("mobilePhone")%></td>
          <td><%="<a href='mailto:"&rs("email")&"'>"&rs("email")&"</a>"%>&nbsp;</td>
          <td><%=left(rs("comment"),30)%>&nbsp;</td>
          <td><%="<a href='viewDetail.asp?id="&rs("id")&"'>����</a>"%></td>
          <td><%="<a href=javascript:del('"&rs("id")&"');>ɾ��</a>"%></td>
        </tr>
        <%rs.movenext
loop%>
      </table></td>
  </tr>
  <tr> 
    <td align="right"> <%				  
''''''''''''''''''''''
if totalrec<>0 then
response.write "�� <font color=#7c0000><b>"&totalrec&"</b></font> �� "
	if currentpage-1 mod 10=0 then
		p=(currentpage-1) \ 10
	else
		p=(currentpage-1) \ 10
	end if
	dim pagelist,pagelistbit
	if currentPage=1 then
		response.write "<font face=webdings color=#ff0000>9</font>   "
	else
		response.write "<a href='?sortid="&sortid&"&page=1' title=��ҳ><font face=webdings>9</font></a>   "
	end if
	if p*10>0 then response.write "<a href='?sortid="&sortid&"&page="&Cstr(p*10)&"' title=��ʮҳ><font face=webdings>7</font></a>   "
	response.write "<b>"
	for ii=p*10+1 to P*10+10
		if ii=currentPage then
			response.write "<font color=#ff0000>"+Cstr(ii)+"</font> "
		else
			response.write "<a href='?sortid="&sortid&"&page="&Cstr(ii)&"'>"+Cstr(ii)+"</a>   "
		end if
		if ii=n then exit for
		'p=p+1
	next
	response.write "</b>"
	if ii<n then response.write "<a href='?sortid="&sortid&"&page="&Cstr(ii)&"' title=��ʮҳ><font face=webdings>8</font></a>   "
	if currentPage=n then
		response.write "<font face=webdings color=#ff0000>:</font>   "
	else
		response.write "<a href='?sortid="&sortid&"&page="&Cstr(n)&"' title=βҳ><font face=webdings>:</font></a>"
	end if
end if
''''''''''''''''''''''''''''''''''''''''''
%> </td>
  </tr>
</table>
</body>
</html>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
