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
   if (! confirm("你真的想删除此记录吗？"))
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
response.write "没有找到"

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
 response.write "目前没有"%>
            <%elseif pagen=1 then%>
            &nbsp;&nbsp; <span >页次：</span>
            <% response.write "第"&pageid&"页"&"/共"&pagen&"页"%>
            &nbsp;&nbsp;&nbsp;&nbsp; 
            <%elseif pageid=1 then%>
            &nbsp;&nbsp; 页次：
            <% response.write "第"&pageid&"页"&"/共"&pagen&"页"%>
            &nbsp;&nbsp;&nbsp;&nbsp; <span >｜</span><a href="newsedlist.asp?searchtxt=<%=searchtxt%>&sortid=<%=sortid%>&pageid=<%response.write (pageid+1)%>">下一页</a> 
            <%elseif abs(pagen)=abs(pageid) then%>
            &nbsp;&nbsp; 页次：
            <% response.write "第"&pageid&"页"&"/共"&pagen&"页"%>
            &nbsp;&nbsp;&nbsp;&nbsp; <a href="newsedlist.asp?searchtxt=<%=searchtxt%>&sortid=<%=sortid%>&pageid=<%response.write (pageid-1)%>">上一页</a><span >｜</span> 
            <%else%>
            &nbsp;&nbsp; 页次：
            <% response.write "第"&pageid&"页"&"/共"&pagen&"页"%>
            &nbsp;&nbsp;&nbsp;&nbsp; <a href="newsedlist.asp?searchtxt=<%=searchtxt%>&sortid=<%=sortid%>&pageid=<%response.write (pageid-1)%>"> 
            上一页</a><span >｜</span><a href="newsedlist.asp?searchtxt=<%=searchtxt%>&sortid=<%=sortid%>&pageid=<%response.write (pageid+1)%>">下一页</a> 
            <%end if%>
        </td>
      </tr>
      <tr bgcolor="#CCCCCC"> 
        <td width="47" height="25" align="center"> ID</td>
        <td width="" align="center"> 新闻标题</td>
        <td width="145" align="center"> 时间</td>
        <td width="65" align="center"> 修改</td>
        <td width="65" align="center"> 审核</td>
        <td width="65" align="center"> 删除</td>
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
            <%if rs("news_tpic")<>"" then response.write "<font color=red>(图)</font>"%>
        </td>
        <td align="center" width="138" bordercolorlight="#D0D0D0" bordercolordark="#FFFFFF" bordercolor="#C0C0C0"> 
        <%=rs("news_uptime")%></td>
        <td align="center" width="65" bordercolorlight="#D0D0D0" bordercolordark="#FFFFFF" bordercolor="#C0C0C0"> 
          <p align="center"> <a href="newsedit.asp?id=<%=rs("id")%>">修改</a> </td>
        <td align="center" style="font-size: 9pt; font-family:宋体; color:#000000" width="65" bordercolorlight="#D0D0D0" bordercolordark="#FFFFFF" bordercolor="#C0C0C0"> 
          <p align="center"> 
            <%if rs("news_level")=1 then%>
            <a style="text-decoration: none" href="chlevel.asp?level=<%=rs("news_level")%>&id=<%=rs("id")%>"> 
            已审核</a> 
            <%else%>
            <a href="chlevel.asp?level=<%=rs("news_level")%>&id=<%=rs("id")%>"><font color="#808080">未审核</font></a>
            <%end if%>
        </td>
        <td width="65" align="center" bordercolor="#C0C0C0" bordercolorlight="#D0D0D0" bordercolordark="#FFFFFF" bgcolor="#FFFFFF"> 
          <p align="center"> <span ><font color="#000000"> 
          
                   <a style="text-decoration: none" href="newsdel.asp?id=<%=rs("id")%>"onclick="return delcon();">删除</a></font></span></a><span ><a style="text-decoration: none" href="dele.asp?id=<%=rs("sortid")%>" > 
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