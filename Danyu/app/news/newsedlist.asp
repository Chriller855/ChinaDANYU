<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!-- #include file="conn.asp" -->
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../index.css" rel="stylesheet" type="text/css">
<script language="javascript">
function delcon()
 { 
   if (! confirm("�������ɾ���˼�¼��"))
   {    return false;
   }
}
</script>
<%
set rs=server.createobject("adodb.recordset")
searchtxt=request("searchtxt")
searchtxt=replace(replace(replace(searchtxt,"'",""),"%",""),"=","")
sortid=clng(request("sortid"))
if sortid=0 then
sql1="where"
else
sql1="WHERE sortid = "&sortid
end if

if searchtxt="" then
sql2=""
else
sql2=" news_title LIKE '%"&searchtxt&"%'"
end if

if searchtxt="" and sortid=0 then
sql1=""
end if

if searchtxt="" or sortid=0 then
sql3=""
else 
sql3=" and "
end if


sql="select * from article "&sql1&sql3&sql2&" order by news_uptime desc "
rs.open sql,conn,1,1
if rs.eof and rs.bof then
response.write "û���ҵ�"

response.end
end if
dim num
dim psize
dim pageid
num=rs.recordcount
rs.pagesize=20
psize=rs.pagesize
pagen=rs.pagecount

pageid=request("pageid")
if pageid="" then
pageid=1
end if
if cint(pageid)>pagen then
pageid=pagen
end if
session("refererString")=request.ServerVariables("SCRIPT_NAME")&"?searchtxt="&searchtxt&"&sortid="&sortid&"&pageid="&pageid

rs.absolutepage=pageid
%>
</head>
<body>
                      <div align="center">
                        <center>

                      
    <table border="1" cellspacing="0" width="98%" bordercolorlight="#D0D0D0" bordercolordark="#D0D0D0" cellpadding="1" style="border-collapse: collapse">
      <tr> 
        <td align="right" bgcolor="#FFFFFF" bordercolor="#FFFFFF" colspan="5" height="22"> 
          <p align="right">
            <%if rs.eof and rs.bof then
 response.write "Ŀǰû��"%>
            <%elseif pagen=1 then%>
            &nbsp;&nbsp; <span >ҳ�Σ�</span>
            <% response.write "��"&pageid&"ҳ"&"/��"&pagen&"ҳ"%>
            &nbsp;&nbsp;&nbsp;&nbsp; 
            <%elseif pageid=1 then%>
            &nbsp;&nbsp; ҳ�Σ�
            <% response.write "��"&pageid&"ҳ"&"/��"&pagen&"ҳ"%>
            &nbsp;&nbsp;&nbsp;&nbsp; <span >��</span><a href="newsedlist.asp?searchtxt=<%=searchtxt%>&sortid=<%=sortid%>&pageid=<%response.write (pageid+1)%>">��һҳ</a> 
            <%elseif abs(pagen)=abs(pageid) then%>
            &nbsp;&nbsp; ҳ�Σ�
            <% response.write "��"&pageid&"ҳ"&"/��"&pagen&"ҳ"%>
            &nbsp;&nbsp;&nbsp;&nbsp; <a href="newsedlist.asp?searchtxt=<%=searchtxt%>&sortid=<%=sortid%>&pageid=<%response.write (pageid-1)%>">��һҳ</a><span >��</span> 
            <%else%>
            &nbsp;&nbsp; ҳ�Σ�
            <% response.write "��"&pageid&"ҳ"&"/��"&pagen&"ҳ"%>
            &nbsp;&nbsp;&nbsp;&nbsp; <a href="newsedlist.asp?searchtxt=<%=searchtxt%>&sortid=<%=sortid%>&pageid=<%response.write (pageid-1)%>"> 
            ��һҳ</a><span >��</span><a href="newsedlist.asp?searchtxt=<%=searchtxt%>&sortid=<%=sortid%>&pageid=<%response.write (pageid+1)%>">��һҳ</a> 
            <%end if%>
        </td>
      </tr>
      <tr bgcolor="#CCCCCC"> 
        <td width="47" height="25" align="center"> ID</td>
        <td width="" align="center"> ���ű���</td>
        <td width="145" align="center"> ʱ��</td>
        <td width="65" align="center"> �޸�</td>
        <td width="65" align="center"> ���</td>
        <td width="65" align="center"> ɾ��</td>
      </tr>
      <% 
dim a
dim b
b=pageid*psize
if num-b < 0 Then 
b=num
else
end if
for a=((pageid-1)*psize+1) to b
%>
      <tr bgcolor="#FFFFFF"> 
        <td height="23" align="center" width="47" bordercolorlight="#D0D0D0" bordercolordark="#FFFFFF" bordercolor="#C0C0C0"> 
          <p align="center"><%=rs("id")%> </td>
        <td align="left" width="510" bordercolorlight="#D0D0D0" bordercolordark="#FFFFFF" bordercolor="#C0C0C0"> 
          <p>
            <%
								response.write "<a href=newsDetailView.asp?id="&rs("id")&" target=_blank>"&rs("news_title")&"</a>"
							%>
            <%if rs("news_tpic")<>"" then response.write "<font color=red>(ͼ)</font>"%>
        </td>
        <td align="center" width="138" bordercolorlight="#D0D0D0" bordercolordark="#FFFFFF" bordercolor="#C0C0C0"> 
        <%=rs("news_uptime")%></td>
        <td align="center" width="65" bordercolorlight="#D0D0D0" bordercolordark="#FFFFFF" bordercolor="#C0C0C0"> 
          <p align="center"> <a href="newsedit.asp?id=<%=rs("id")%>">�޸�</a> </td>
        <td align="center" style="font-size: 9pt; font-family:����; color:#000000" width="65" bordercolorlight="#D0D0D0" bordercolordark="#FFFFFF" bordercolor="#C0C0C0"> 
          <p align="center"> 
            <%if rs("news_level")=1 then%>
            <a style="text-decoration: none" href="chlevel.asp?level=<%=rs("news_level")%>&id=<%=rs("id")%>"> 
            �����</a> 
            <%else%>
            <a href="chlevel.asp?level=<%=rs("news_level")%>&id=<%=rs("id")%>"><font color="#808080">δ���</font></a>
            <%end if%>
        </td>
        <td width="65" align="center" bordercolor="#C0C0C0" bordercolorlight="#D0D0D0" bordercolordark="#FFFFFF" bgcolor="#FFFFFF"> 
          <p align="center"> <span ><font color="#000000"> 
          
                   <a style="text-decoration: none" href="newsdel.asp?id=<%=rs("id")%>"onclick="return delcon();">ɾ��</a></font></span></a><span ><a style="text-decoration: none" href="dele.asp?id=<%=rs("sortid")%>" > 
            </a> 

            
            </span> </td>
      </tr>
      <%rs.movenext 
next

 %>
      <span > </span> 
    </table>
                        </center>
</div>
<p align="center"></p>

<%rs.close 
set rs=nothing
conn.close
set conn=nothing
%>                    
</body>
</html>