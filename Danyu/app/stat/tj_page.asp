<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!-- #include file="inc_config.asp" -->
<%
'Ȩ�޼��
if session.Contents("master")=false and whatcan<2 then Response.Redirect "help.asp?id=004&error=��û�в鿴����ҳ��ͳ�Ƶ�Ȩ�ޡ�"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<title><%=countname%>��������ҳ��ͳ��</title>
<link rel="stylesheet" type="text/css" href="../../index.css">
</head>
<body topmargin=5 rightmargin=0 leftmargin=0 vlink=#000000>
<%
Set rs = Server.CreateObject("ADODB.Recordset")
%>

<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr height="30">
    <td width="498">������ҳ�漰������<br>
        <table border="0" cellpadding="0" cellspacing="0" width="550">
          <tr height="9">
            <td></td>
          </tr>
          <tr height="10">
            <td></td>
            <td width="230"><img src="images/tu_back_up.gif"></td>
          </tr>
          <%

sql="select vpage,count(id) as allpage from view group by vpage order by count(id) DESC"
rs.Open sql,conn,1,1

maxallpage=0
sumallpage=0
do while not rs.EOF
	if clng(rs("allpage"))>maxallpage then maxallpage=clng(rs("allpage"))
	sumallpage=sumallpage+clng(rs("allpage"))
	rs.MoveNext
loop
	'��ֹ����Ϊ�������
	if maxallpage=0 then maxallpage=1
	if sumallpage=0 then sumallpage=1

rs.MoveFirst 

j=0
do while not rs.EOF
thepage=rs("vpage")
vallpage=rs("allpage")
	thelen=len(thepage)
	if thelen =0 then
		thepage="main.asp"
		svpage="ͨ���ղػ�ֱ��������ַ����"
	end if
	if thelen <= 53 and thelen > 0 then
		svpage=thepage
	end if
	if thelen >= 54 then
		svpage=left(thepage,51) & "..."
	end if
%>
          <tr>
            <td align=right><a href="<%=thepage%>" target="_blank"
	title="<%=thepage & vbcrlf%>����<%=vallpage%>�Σ�<%
	'����������İٷ�������ȷ��С����1λ��С�������ǰ�����ĸ0
	lsbf=int(vallpage*1000/sumallpage)/10
	if lsbf<1 then lsbf="0" & lsbf
	Response.Write lsbf
	%>%"><%
	tmp="http://"&request.ServerVariables("SERVER_NAME")
	if request.ServerVariables("SERVER_PORT")<>"80" then tmp=tmp&":"&request.ServerVariables("SERVER_PORT")
	svpage=replace(svpage,tmp&"/stat.asp?s=","")
	svpage=replace(svpage,tmp&"/default.htm","")
	svpage=replace(svpage,tmp&"/default.asp","")
	svpage=replace(svpage,tmp&"/index.htm","")
	svpage=replace(svpage,tmp&"/index.asp","")
	svpage=replace(svpage,tmp,"")
	if svpage="" then svpage="[��ҳ]"
	'response.write tmp
	%><%=svpage%></a>&nbsp;</td>
            <td width="230" background="images/tu_back_2.gif" align=left><img style="BORDER-left: #000000 1px solid;" src="images/tu.gif"
	width="<%=(vallpage/maxallpage)*183%>" height="9"
	alt="<%=thepage%>�꣬����<%=vallpage%>�Σ�<%
	'����������İٷ�������ȷ��С����1λ��С�������ǰ�����ĸ0
	lsbf=int(vallpage*1000/sumallpage)/10
	if lsbf<1 then lsbf="0" & lsbf
	Response.Write lsbf
	%>%"> <%=vallpage%></td>
          </tr>
          <%
rs.MoveNext
	j=j+1
	'�����¼����40�������˳�
	if j=40 then exit do
loop
%>
          <tr height="10">
            <td></td>
            <td width="230"><img src="images/tu_back_down.gif"></td>
          </tr>
          <tr height="5">
            <td colspan=29></td>
          </tr>
      </table></td>
  </tr>
</table>
������ָͼ����������ַ���Կ�����Ӧ�ķ�������Ҫ�õ�ĳһʱ�ε�ͳ����Ϣ����ʹ���Զ���ͳ�ơ�
<%
rs.Close

set rs=nothing
conn.Close 
set conn=nothing
%>
</body>
</html>
<%
'����ָ�����ڵķ�����
function vdaycon(theday)
    theday=cdate(theday)
    thetday=cdate(theday+1)
    tmprs=conn.execute("Select count(id) as vdaycon from view where" & _
		" vtime>=datevalue('" & theday & "') and vtime<=datevalue('" & thetday & "')")
    vdaycon=tmprs("vdaycon")
	if isnull(vdaycon) then vdaycon=0
end function
%>