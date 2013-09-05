<!-- #include file="../user/chkpas.asp" -->
<!-- #include file="inc_config.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<title><%=countname%></title>
<link rel="stylesheet" type="text/css" href="../../index.css">
</head>
<%if yesleft=1 then%>
<frameset rows="*" cols="86,*" >
  <frame name="countleft" scrolling="NO" noresize src="list.asp">
  <%if Request("main")="master" then%>
  <frame name="countmain" src="master.asp">
  <%else%>
  <frame name="countmain" src="main.asp">
  <%end if%>
</frameset>
<%else%>
<frameset cols="*" frameborder="NO" border="0" framespacing="0" rows="*"> 
<%if Request("main")="master" then%>
  <frame name="countmain" src="master.asp">
<%else%>
  <frame name="countmain" src="main.asp">
<%end if%>
</frameset>
<%end if%>
<noframes>
<body>
<a href="main.asp">您的浏览器不支持框架，只好点这里进入了。</a>
</body>
</noframes> 
</html>