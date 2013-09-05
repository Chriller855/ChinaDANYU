<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!-- #include file="conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title></title>
<link href="../index.css" rel="stylesheet" type="text/css">
</head>
<body>

<table width="100%" border="0" cellpadding="5" cellspacing="1">
  <tr>
    <td> <span style="font-size: 14px">新闻管理系统</span> -&gt; <font color="#000000"><strong>内容修改</strong></font></td>
  </tr>
  <tr>
    <td><br>
      <%iUrl="newsEdList.asp"%><!--#include file="listTitlesSortTable.asp" --></td>
  </tr>
  </table>
<br/><br/>
<p style="font-size:14px;color:#F00;"><strong>注意：</strong>修改招聘信息<a href="newsEdList.asp?sortid=84" target="_self">请点击这里</a><br>
(删除招聘信息请修改其内容为：&ldquo;暂无招聘，敬请关注！&rdquo;)</p>
</body>

</html>