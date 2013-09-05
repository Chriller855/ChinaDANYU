<!-- #include file="inc_config.asp" -->

<%
' 网页立即超时，防止漏统计
Response.Expires = 0

' 被访问页面

if isNull(session("vpage")) then session("vpage")=""
if session("vpage")="" then
	vpage=Request.ServerVariables("HTTP_REFERER")
else
	vpage=session("vpage")
	session("vpage")=""
end if


if right(vpage,1)="/" then vpage=left(vpage,len(vpage)-1)

'检查是否强制页面
onlylen1=len(onlythesite1)
onlylen2=len(onlythesite2)

canstat=false
if (onlylen1=0 and onlylen2=0) then canstat=true
if onlylen1=0 and onlylen2<>0 and left(vpage,onlylen2)=onlythesite2 then canstat=true
if onlylen1<>0 and onlylen2=0 and left(vpage,onlylen1)=onlythesite1 then canstat=true
if onlylen1<>0 and onlylen2<>0 and (left(vpage,onlylen1)=onlythesite1 or left(vpage,onlylen2)=onlythesite2) then canstat=true

if canstat=true then

'****************** 创建数据对象 ******************


Set rs = Server.CreateObject("ADODB.Recordset")

'****************** 记录相关数据 ******************

' IP
vip=Request.ServerVariables("Remote_Addr")
' 来路
if isNull(session("fromWhere")) then session("fromWhere")=""
if session("fromWhere")="" then
vcome=Request("referer")
else
vcome=session("fromWhere")
session("fromWhere")=""
end if
if right(vcome,1)="/" then vcome=left(vcome,len(vcome)-1)

' 被访问页面 （已提前检查）
' 时间
vyear=year(now())
vmonth=month(now())
vday=day(now())
vhour=hour(now())
vtime=now()
vweek=weekday(now())
' 客户端软件使用情况
thesoft=Request.ServerVariables("HTTP_USER_AGENT")
' 浏览器
if instr(thesoft,"NetCaptor") then
	vsoft="NetCaptor"
elseif instr(thesoft,"MSIE 6") then
	vsoft="MSIE 7.x"
elseif instr(thesoft,"MSIE 6") then
	vsoft="MSIE 6.x"
elseif instr(thesoft,"MSIE 5") then
	vsoft="MSIE 5.x"
elseif instr(thesoft,"MSIE 4") then
	vsoft="MSIE 4.x"
elseif instr(thesoft,"Netscape") then
	vsoft="Netscape"
elseif instr(thesoft,"Opera") then
	vsoft="Opera"
else
	vsoft="Other"
end if
' 操作系统
if instr(thesoft,"Windows NT 5.0") then
	vOS="Win 2000"
elseif instr(thesoft,"Windows NT 5.1") then
	vOs="Win XP"
elseif instr(thesoft,"Windows Vista") then
	vOs="Win Vista"
elseif instr(thesoft,"Windows NT") then
	vOs="Win NT"
elseif instr(thesoft,"Windows 9") then
	vOs="Win 9x"
elseif instr(thesoft,"unix") or instr(thesoft,"linux") or instr(thesoft,"SunOS") or instr(thesoft,"BSD") then
	vOs="类Unix"
elseif instr(thesoft,"Mac") then
	vOs="Mac"
else
	vOs="Other"
end if
'屏幕宽度
vwidth=Request("screenwidth")


' 访问者所在地区
' 从独立数据库读取IP信息
set connip=server.createobject("adodb.connection")
DBPath = Server.MapPath(ipconnpath)
connip.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath

ipnow=cip(vip)

Set rsip = Server.CreateObject("ADODB.Recordset")
sql="select * from ip where (onip<=" & ipnow & " and offip>=" & ipnow & ")"
rsip.Open sql,connip,1,1

if rsip.EOF then
	vwhere="未知"
	vwheref=""
else
	vwhere=rsip("addj")
	vwheref=rsip("addf")
end if
rsip.Close
set rsip=nothing
connip.Close
set connip=nothing

'************** 检查是否属于刷新 ****************
if is_ipcheck=true then
	is_echo=cbool(Request.Cookies(mNameEn & "echo")("lao"))
	if is_echo <> true then
	Response.Cookies(mNameEn & "echo")("lao")=true
	Response.Cookies(mNameEn & "echo").Expires=now()+0.01
	end if
end if

'***************** 统计在线人数 *****************
if is_online then
	sql="select vip from view where vtime >= dateadd('n',-20,now()) group by vip"
	rs.Open sql,conn,1,1
	vonline=rs.RecordCount
	rs.Close
end if

'****************** 写主数据库 ******************
if (is_online=true and is_echo <> true) or is_online=flase then
sql="select top 1 * from view"
rs.Open sql,conn,3,2
rs.AddNew
rs("vyear")=vyear
rs("vmonth")=vmonth
rs("vday")=vday
rs("vhour")=vhour
rs("vtime")=vtime
rs("vweek")=vweek
rs("vip")=vip
rs("vwhere")=vwhere
rs("vwheref")=vwheref
rs("vcome")=vcome
rs("vpage")=vpage
rs("vsoft")=vsoft
rs("vos")=vos
rs("vwidth")=vwidth
rs.Update 
rs.Close
end if

'****************** 读写简数据库 ******************
sql="select * from vjian"
rs.Open sql,conn,3,3
' 写入数据
if (is_online=true and is_echo <> true) or is_online=flase then
  if rs.EOF then
	rs.AddNew
	rs("today")=1
	rs("yesterday")=0
	rs("vdate")=date()
	rs("vtop")=1
	rs("starttime")=now()
	rs.Update
  else
	on error resume next
	olddate=cdate(rs("vdate"))
	datecha=date()-olddate
	'上一条记录是今天的
	select case datecha
	case 0
		rs("today")=rs("today")+1
	case 1
		thetoday=rs("today")
		rs("yesterday")=thetoday
		rs("today")=1
		rs("vdate")=date()
	case else
		rs("today")=1
		rs("yesterday")=0
		rs("vdate")=date()
	end select
	
	rs("vtop")=rs("vtop")+1
	rs.Update
  end if
end if

'读出数据
	vtoday=rs("today")
	vyesterday=rs("yesterday")
	vtop=rs("vtop")
	rs.Close
	'加上初始值
	vtop=vtop + old_count

'************ 读写COOKIE，得到用户浏览量 **********

	vuser=cint(Request.Cookies(mNameEn)("lao"))+1
	Response.Cookies(mNameEn)("lao")=vuser
	Response.Cookies(mNameEn).Expires=date()+CookieExpires


'****************** 关闭数据对象 ******************
set rs=nothing
conn.Close 
set conn=nothing

'******************** 写出HTML ********************

'程序及图像文件路径
theurl="http://" & Request.ServerVariables("http_host") & finddir(Request.ServerVariables("url"))

'获取输出要求
style=Request("style")

'根据要求输出
select case style
case "counter"	'LOGO
	outstr="<table width='88' border='0' cellspacing='0' cellpadding='0' height='31' background='" & theurl & "images/stat_counter.gif'><tr><td height='5' width='24'></td><td height='5' width='57'></td><td height='5' width='7'></td></tr><tr><td height='16'></td><td height='16' align='center' valign='top'><marquee behavior='loop' scrollDelay='100' scrollAmount='3' style='font-size: 12px; line-height=15px'><a href='" & theurl & "index.asp' target='_blank' style='color: #ffffff; text-decoration: none'>"
	outstr=outstr & "<font face='Arial, Verdana, san-serif' color='#407526'>总访问量: " & vtop & " &nbsp;今日访问: " & vtoday & " &nbsp;昨日访问: " & vyesterday & " &nbsp;您的访量: "
if is_online then
	outstr=outstr & vuser & " &nbsp;在线人数: " & vonline
end if
	outstr=outstr & "</font>"
	outstr=outstr & "</a></marquee></td><td height='16'></td></tr><tr><td height='10'></td><td height='10'></td><td height='10'></td></tr></table>"

case "icon"		'ICON
	outstr="<a href='" & theurl & "index.asp' title='" & countname &  _
		"\n总量: " & vtop & _
		"\n今日: " & vtoday & "\n昨日: " & vyesterday & _
		"\n您的: " & vuser 
	if is_online then
		outstr=outstr & "\n在线: " & vonline
	end if
	outstr=outstr & "' target='_blank'><img border='0' src='" & theurl & _
	"images/stat_icon.gif'></a>"

case "flash"	'FLASH
	outstr="<embed src='" & theurl & "count.swf?cgi=" & theurl & "flash.asp' type='application/x-shockwave-flash' width='" & FlashWidth & "' height='" & FlashHeight & "'></embed>"

case "text"		'TEXT
	outstr=vtop
end select

'输出
Response.Write "document.write(" & chr(34) & outstr & chr(34) & ")"


else   '检查是否非法调用
	Response.Write "document.write(" & chr(34) & "非法调用" & chr(34) & ")"
end if

'****************** 自定义函数 ********************

function cip(sip)
	tip=cstr(sip)
	sip1=left(tip,cint(instr(tip,".")-1))
	tip=mid(tip,cint(instr(tip,".")+1))
	sip2=left(tip,cint(instr(tip,".")-1))
	tip=mid(tip,cint(instr(tip,".")+1))
	sip3=left(tip,cint(instr(tip,".")-1))
	sip4=mid(tip,cint(instr(tip,".")+1))
	cip=cint(sip1)*256*256*256+cint(sip2)*256*256+cint(sip3)*256+cint(sip4)
end function

Function finddir(filepath)
	finddir=""
	for i=1 to len(filepath)
	if left(right(filepath,i),1)="/" or left(right(filepath,i),1)="\" then
	  abc=i
	  exit for
	end if
	next
	if abc <> 1 then
	finddir=left(filepath,len(filepath)-abc+1)
	end if
end Function
%>