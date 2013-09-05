<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>在线简历</title>
<!--<link rel="stylesheet" type="text/css" href="css/base.css">-->
<link rel="stylesheet" type="text/css" href="css/atsunan.css">
<link rel="stylesheet" type="text/css" href="css/join.css">
<link rel="stylesheet" type="text/css" href="index.css">
<script type="text/javascript" src="Scripts/prototype.js" ></script>
<script type="text/javascript" src="Scripts/join_current.js"></script>
<script src="Scripts/AC_RunActiveContent.js" type="text/javascript"></script>
      <script language="JavaScript">
<!--
function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_validateForm() { //v4.0
  var i,p,q,nm,test,num,min,max,errors='',args=MM_validateForm.arguments;
  for (i=0; i<(args.length-2); i+=3) { test=args[i+2]; val=MM_findObj(args[i]);
    if (val) { nm=args[i+1]; if ((val=val.value)!="") {
      if (test.indexOf('isEmail')!=-1) { p=val.indexOf('@');
        if (p<1 || p==(val.length-1)) errors+='- '+nm+' 必须是一个合法的email地址.\n';
      } else if (test!='R') {
        if (isNaN(val)) errors+='- '+nm+' 必须是一数字.\n';
        if (test.indexOf('inRange') != -1) { p=test.indexOf(':');
          min=test.substring(8,p); max=test.substring(p+1);
          if (parseFloat(val)<parseFloat(min) || parseFloat(max)<parseFloat(val)) errors+='- '+nm+' 必须是数字而且在'+min+' 和 '+max+'之间。\n';
    } } } else if (test.charAt(0) == 'R') errors += '- '+nm+' 是必须填写的项目。\n'; }
  }if (errors) alert('由于下面的错误，您的表单不能提交:\n'+errors);
  document.MM_returnValue = (errors == '');
}
//-->
          </script>   

</head>

<body topmargin="0" leftmargin="0" bgcolor="#F5F5F5">

<div align="center">
	<table border="0" cellpadding="0" cellspacing="0" id="table1" width="940" height="80" bgcolor="#FFFFFF">
		<tr>
			<td>
			  <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="940" height="80">
          <param name="movie" value="flash/menu.swf?id=6">
		  <param name="wmode" value="transparent">
          <param name="quality" value="high">
          <embed src="flash/menu.swf?id=6" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="940" height="80">
          </embed>
          </object> 
			</td>
		</tr>
	</table>
</div>
<div align="center">
	<table border="0" cellpadding="0" cellspacing="0" width="940" id="table2" bgcolor="#FFFFFF">
		<tr>
			<td><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=9,0,28,0" width="940" height="184">
			  <param name="movie" value="flash/xtl01.swf">
			  <param name="quality" value="high">
			  <param name="wmode" value="opaque">
			  <embed src="flash/xtl01.swf" quality="high" wmode="opaque" pluginspage="http://www.adobe.com/shockwave/download/download.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="940" height="184"></embed>
		    </object></td>
		</tr>
	</table>
</div>
<div align="center">
	<table border="0" cellpadding="0" cellspacing="0" width="940" id="table3" bgcolor="#FFFFFF" style="border-bottom: 3px solid #E5E5E5">
		<tr>
			<td height="40" colspan="4">　</td>
		</tr>
		<tr>
			<td width="24" valign="top">　</td>
			<td width="197" valign="top">
			<div id="mainLeft">
    <!--左侧菜单开始-->
    <div id="menuOut">
      <div id="menuPor">
<div id="menu"><a href="join01.asp" id="m1"></a> <a href="join02.asp" id="m2"></a><a href="join03.asp" id="m3"></a> 
</div>
</div>
    </div>
			</td>
			<td width="28" valign="top">　</td>
			<td width="691" valign="top">
			<div align="center">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table5">
					<tr>
						<td colspan="2">
						<img border="0" src="images/tit0603.gif" width="691" height="47"></td>
					</tr>
					<tr>
						<td width="670">
<%
session.timeout=10
if request("action")="add" then
%>
      <!--#include file="app/visitApp/conn.asp" -->
      <%
	if datediff("s",session("updateTime"),now())<180 then
		response.write "<br><br><p align=center><font color=#6D2401 style='font-size:14px;'>请勿重复提交!或过3分钟后重新提交活动报名申请！！ </font><br><br><a href=javascript:history.back();><font color=#6D2401 style='font-size:14px;'>[返回]</font></a></p>"
		conn.close
		set conn=nothing
	else
		ierr=""
		'if request("visitApp_j") ="" then ierr=ierr&"意向房型/Intended House 是必须填写的项目。<BR>"
		if ierr<>"" then
			response.write "<br><br>"&ierr
			response.write "<p align=center><font color=#6D2401 style='font-size:14px;'>请返回重填！！ </font><br><br><a href='javascript:history.go(-1);'><font color=#6D2401 style='font-size:14px;'>[返回]</font></a></p>"
		else
			set rs = createObject("adodb.recordset")
			sql="select * from visitApp"
			rs.open sql,conn,3,4
			rs.addnew
			rs("postCode") = server.HTMLEncode(request("postCode"))
			rs("memberName") = server.HTMLEncode(request("memberName"))
			rs("address") = server.HTMLEncode(request("address"))
			rs("comeDate") = server.HTMLEncode(request("comeDate"))
			rs("ok") = server.HTMLEncode(request("ok"))
			rs("ok2") = server.HTMLEncode(request("ok2"))
			rs("telphone") = server.HTMLEncode(request("telphone"))
			rs("mobilePhone") = server.HTMLEncode(request("mobilePhone"))
			rs("email") = server.HTMLEncode(request("email"))
			rs("comment") = server.HTMLEncode(request("comment"))
			rs("dateAndTime") = now()
			rs("ip")=request.ServerVariables("REMOTE_ADDR")
			
			rs.updatebatch
			session("updateTime")=now()
			rs.close
			set rs=nothing
			conn.close
			set conn=nothing
			
			response.write "<br><br><p align=center><font color=#525252 style='font-size:14px;'>表格提交成功，您的申请在审核中,我们将尽快与您联系！</font><br><br><a href="
			response.write request.ServerVariables("SCRIPT_NAME")&"><font color=#525252 style='font-size:14px;'>返回</font></a></p>"
		end if
	end if
else
if session("updateTime")="" then session("updateTime")=cdate("1900-01-01")
%>
          
 
                        
                        <table width="600" border="0" cellpadding="2" cellspacing="0">
  <form action="join03.asp?action=add" method="POST" name="userForm" id="userForm" onSubmit="MM_validateForm('ok2','工作经验','R','ok','学历','R','comeDate','应聘职位','R','memberName','姓名','R','telphone','电话','R','telphone','电话','inRange','mobilePhone','手机','R','mobilePhone','手机','inRange','email','电子邮件','R','email','电子邮件','isEmail','address','年龄','R','postCode','称呼');return document.MM_returnValue">
    <tr> 
      <td>
        <p style="margin-bottom: 0; margin-top:0"> 
		<img border="0" src="images/fk.gif" width="364" height="30"></p>
      </td>
    </tr>
    <tr> 
      <td width="301"></td>
    </tr>
    <tr> 
      <td height="34"><table border="0" cellspacing="0" cellpadding="0" width="100%">
          <tr> 
            <td width="15%">称&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;呼：</td>                                                                
            <td width="85%">
                    <input name="postCode" type="radio" class="trans" value="先生" checked>先生&nbsp;&nbsp;&nbsp;&nbsp;                     
                    <input name="postCode" type="radio" class="trans" value="女士">女士&nbsp;&nbsp;&nbsp;&nbsp;                                                           
                      <input name="postCode" type="radio" class="trans" value="小姐">小姐&nbsp;</td>
          </tr>
        </table>
      </td>
    </tr>
    <tr> 
      <td height="34"><table border="0" cellspacing="0" cellpadding="0" width="100%">
          <tr> 
            <td width="15%">姓&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;名：</td>                                                                  
            <td width="31%">
            <input type="text" name="memberName" size="20" style="border: 1px solid #cccccc; padding: 0; background-color:#F5F5F5"></td>
            <td width="15%">年&nbsp;&nbsp;&nbsp;&nbsp;龄：</td>                                                                
            <td width="39%"><input name="address" type="text" style="border: 1 solid #cccccc; padding: 0; background-color:#F5F5F5" size="20"></td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td height="34"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="15%">应&nbsp;聘&nbsp;职&nbsp;位：</td>                                             
            <td width="31%">
            <input type="text" name="comeDate" size="20" style="border: 1 solid #cccccc; padding: 0; background-color:#F5F5F5"></td>
            <td width="15%">学&nbsp;&nbsp;&nbsp;&nbsp;历：</td>                                                                
            <td width="39%"><input name="ok" type="text" style="border: 1 solid #cccccc; padding: 0; background-color:#F5F5F5" size="20"></td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td>&nbsp;
      </td>
    </tr>
    <tr> 
      <td><table border="0" cellspacing="0" cellpadding="0" width="100%">
          <tr> 
            <td width="15%" valign="top">
              <p>工&nbsp;作&nbsp;经&nbsp;验：</p>                                          
            </td>
            <td width="85%">
<select name="ok2"  style="width: 113; height: 23"><option value=''>---请选择---</option>
<option style="width:100px" value='在读学生'>在读学生</option>
<option style="width:100px" value='应届毕业生'>应届毕业生</option>
<option style="width:100px" value='一年以上'>一年以上</option>
<option style="width:100px" value='二年以上'>二年以上</option>
<option style="width:100px" value='三年以上'>三年以上</option>
<option style="width:100px" value='五年以上'>五年以上</option>
<option style="width:100px" value='八年以上'>八年以上</option>
<option style="width:100px" value='十年以上'>十年以上</option></select>            </td> 
          </tr>
        </table>
      </td>
    </tr>
    <tr> 
      <td>　</td>
    </tr>
    <tr> 
      <td height="34"><table border="0" cellspacing="0" cellpadding="0" width="100%">
          <tr> 
            <td width="15%">电&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;话：</td>                                                                  
            <td width="31%">
            <input name="telphone" type="text" style="border: 1 solid #cccccc; padding: 0; background-color:#F5F5F5" size="20"></td>
            <td width="15%">手&nbsp;&nbsp;&nbsp;&nbsp;机：</td>                                                                
            <td width="39%"><input name="mobilePhone" type="text" style="border: 1 solid #cccccc; padding: 0; background-color:#F5F5F5" size="20"></td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td height="34"><table border="0" cellspacing="0" cellpadding="0" width="100%">
          <tr> 
            <td width="15%">电&nbsp;子&nbsp;邮&nbsp;件：</td>
            <td width="85%"><input name="email" type="text" style="border: 1 solid #cccccc; padding: 0; background-color:#F5F5F5" size="40">
            </td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td></td>
    </tr>
    <tr>
      <td height="28">
        <p style="margin-top: 6">个人工作简历：</td>
    </tr>
    <tr> 
      <td>
      <textarea cols="81" rows=5 name="comment" wrap="soft" style="border: 1 solid #cccccc; padding: 5px; background-color:#F5F5F5"></textarea></td>
    </tr>
    <tr> 
      <td height="15">
        <p style="margin-left:150px; margin-top:15px;"><input type="submit" name="Submit2" value="--- 提交个人求职信息 ---" style="background-color: #CA0808; color: #FFFFFF; border: 1 solid #CA0808"></p>
      </td>
    </tr>
  </form>
</table>


						</td>
						
					</tr>
					<tr>
						<td>　</td>
					</tr>
				</table>
			</div>
			</td>
		</tr>
		<tr>
			<td colspan="4">　</td>
		</tr>
	</table>
</div>
<div align="center">
	<table border="0" cellpadding="0" cellspacing="0" width="940" id="table4" bgcolor="#FFFFFF">
		<tr>
			<td>
			<!--#include file="bottom.htm" -->
			</td>
		</tr>
	</table>
</div>

</body>
<%end if%>
</html>