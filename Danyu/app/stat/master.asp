<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!-- #include file="inc_config.asp" -->
<%

'****************** 创建数据对象 ******************
Set rs = Server.CreateObject("ADODB.Recordset")

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<title><%=countname%>－管理员控制台</title>
<link rel="stylesheet" type="text/css" href="../../index.css">
</head>
<body topmargin=5 rightmargin=0 leftmargin=0 vlink=#000000>
<br>
&nbsp; <font style="FONT-SIZE: 16px" 
     >&nbsp;</font>管理员控制台<br>
<%if isUpgrade() or (isupgrade12()=false) or isupgrade13()=false then%>
<ul>
  <li style="MARGIN-TOP:	 15px; MARGIN-BOTTOM: 5px; MARGIN-LEFT: 50px"> <A href="upgrade1.asp">数据库结构升级</a>
      <p>如果原来是1.0版，请确认已经上传了升级相关的文件。本操作删除1.0版数据库文件中的IP库，并增加1.1以前版本不具备的的屏幕宽度字段，和1.2版本不具备的备份标记字段及开始统计日期字段。
          <%end if%>
      </p>
  <li style="MARGIN-TOP:	 15px; MARGIN-BOTTOM: 5px; MARGIN-LEFT: 50px"> 数据备份和清理
      <p>统计器使用一段时间以后，主访问记录数据库会变得很大，这不但占用了大量的网站空间，还使统计器的运行效率大大降低，所以应该定期清理访问记录数据库。
      <p>为了保证被清理的数据不至于完全丢失，应该在清理前先使用备份工具备份要清理的数据，现在的清理工具不会清理没有备份过的记录。
      <p><A href="bak_step1.asp" target=_blank>数据备份&gt;&gt;&gt;</a> &nbsp;&nbsp;<A href="clear_step1.asp" target=_blank>数据清理&gt;&gt;&gt;</a>
  <li style="MARGIN-TOP:	 15px; MARGIN-BOTTOM: 5px; MARGIN-LEFT: 50px"> <A href="update.asp?type=go&amp;start=yes" >用新的IP数据库更新所有已有数据</a>
      <p>本操作将占用大量资源，建议现有记录超过10000的朋友不要更新了。执行更新时屏幕将会不断刷新，这是程序在执行，大约10分钟可以更新10000条信息。更新可以随时中止，并可到本控制台继续执行。
      <p>本操作仅当更换了IP数据库时使用，无需进行本操作。
            <%if isupdate() then%>
      </p>
    <li style="MARGIN-TOP:	 15px; MARGIN-BOTTOM: 5px; MARGIN-LEFT: 50px"> <A href="update.asp?type=go">继续上次未完成的更新</a>
        <p>系统检查到您上一次对主数据库的更新尚未完成，点击上面的连接可继续进行。
            <%end if%>
        </p>
    <li style="MARGIN-TOP:	 15px; MARGIN-BOTTOM: 5px; MARGIN-LEFT: 50px"> <A href="logout.asp" target=_top>退出控制台</a>
        <p>为防止您的管理员身份被别人利用，请离开电脑时关闭本窗口或者点击上面的退出连接。</p>
    </li>
</ul>
</body>
</html>
<%

'****************** 关闭数据对象 ******************
set rs=nothing
conn.Close 
set conn=nothing

'****************** 自定义函数 ********************
function isUpgrade()
	isUpgrade=true
	on error resume next
	sql="select * from ip"
	conn.execute(sql)
	if err<>0 then isUpgrade=false
end function

function isUpgrade12()
	isUpgrade12=true
	on error resume next
	sql="select vwidth from view"
	conn.execute(sql)
	if err<>0 then isUpgrade12=false
end function

function isUpgrade13()
	isUpgrade13=true
	on error resume next
	sql="select bakdays from view"
	conn.execute(sql)
	if err<>0 then isUpgrade13=false
end function

function isUpdate()
	isUpdate=true
	on error resume next
	sql="select * from view where up_date=0"
	conn.execute(sql)
	if err<>0 then isUpdate=false
end function
%>