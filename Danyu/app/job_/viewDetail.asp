<!-- www.00ok.net 版权所有 -->
<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!--#include file="conn.asp"-->
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../../index.css" rel="stylesheet" type="text/css">
</head>
<body>
<%
set rs = createObject("adodb.recordset")
sql="select * from job where id="&request("id")
rs.open sql,conn,1,1
%>
<table width="94%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td align="right"><div align="left">应聘职位：<%=rs("j01")%></div></td>
    <td align="right"><div align="center">希望薪资待遇：<%=rs("j02")%></div></td>
    <td align="right">登记时间：<%=rs("j03")%></td>
  </tr>
</table>
<table width="94%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolordark="#FFFFFF">
  <tr bgcolor="#D5D5D5"> 
    <td colspan="5" nowrap> <strong>一. 个人资料</strong></td>
  </tr>
  <tr> 
    <td colspan="3" nowrap> 姓名：<%=rs("j04")%></td>
    <td width="36%" nowrap>性别：<%=rs("j05")%></td>
    <td width="29%" nowrap>年龄：<%=rs("j06")%></td>
  </tr>
  <tr> 
    <td colspan="3" nowrap>籍贯：<%=rs("j07")%></td>
    <td width="36%" nowrap>婚姻状况：<%=rs("j08")%></td>
    <td width="29%" nowrap>有无子女：<%=rs("j09")%></td>
  </tr>
  <tr> 
    <td colspan="3" nowrap>政治面貌：<%=rs("j10")%></td>
    <td width="36%" nowrap>最高学历：<%=rs("j11")%></td>
    <td width="29%" nowrap>身高：<%=rs("j12")%></td>
  </tr>
  <tr> 
    <td colspan="5" nowrap>现在住址：<%=rs("j13")%></td>
  </tr>
  <tr> 
    <td colspan="3" nowrap>个人手机：<%=rs("j14")%></td>
    <td width="36%" nowrap>E-mail：<%="<a href='mailto:"&rs("j15")&"'>"&rs("j15")&"</a>"%></td>
    <td width="29%" nowrap>何时可以开始工作：<%=rs("j16")%></td>
  </tr>
  <tr bgcolor="#F6F6F6"> 
    <td colspan="5" nowrap>紧急情况联系人</td>
  </tr>
  <tr bgcolor="#F6F6F6"> 
    <td colspan="2" nowrap>姓名：<%=rs("j17")%></td>
    <td width="13%" nowrap>关系：<%=rs("j18")%></td>
    <td width="36%" nowrap>地址：<%=rs("j19")%></td>
    <td width="29%" nowrap>电话：<%=rs("j20")%></td>
  </tr>
  <tr bgcolor="#CCCCCC"> 
    <td colspan="5" nowrap><strong>二. 专业、教育程度</strong></td>
  </tr>
  <tr bgcolor="#EFEFEF"> 
    <td width="11%" nowrap><div align="center">自年月</div></td>
    <td width="11%" nowrap><div align="center">至年月</div></td>
    <td nowrap><div align="center">学历类别</div></td>
    <td width="36%" nowrap><div align="center">毕业学校</div></td>
    <td width="29%" nowrap><div align="center">学习专业</div></td>
  </tr>
  <tr> 
    <td nowrap> <div align="center"><%=rs("b01")%></div></td>
    <td nowrap> <div align="center"><%=rs("b02")%></div></td>
    <td nowrap> <div align="center"><%=rs("b03")%></div></td>
    <td width="36%" nowrap> <div align="center"><%=rs("b04")%></div></td>
    <td width="29%" nowrap> <div align="center"><%=rs("b04")%></div></td>
  </tr>
  <tr> 
    <td nowrap> <div align="center"><%=rs("c01")%></div></td>
    <td nowrap> <div align="center"><%=rs("c02")%></div></td>
    <td nowrap> <div align="center"><%=rs("c03")%></div></td>
    <td width="36%" nowrap> <div align="center"><%=rs("c04")%></div></td>
    <td width="29%" nowrap> <div align="center"><%=rs("c05")%></div></td>
  </tr>
  <tr> 
    <td colspan="3" nowrap>个人持有其它资质证书：<%=rs("j21")%></td>
    <td width="36%" nowrap>语言能力（外语） 语种：<%=rs("j22")%></td>
    <td width="29%" nowrap>掌握程度：<%=rs("j23")%></td>
  </tr>
  <tr> 
    <td colspan="5" nowrap>可熟练操作的应用软件：</td>
  </tr>
  <tr> 
    <td colspan="5" nowrap bgcolor="#CCCCCC"><strong>三. 工作经验</strong></td>
  </tr>
  <tr> 
    <td colspan="5" nowrap><table width="100%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolordark="#FFFFFF">
        <tr> 
          <td width="10%"> <div align="center">自年月</div></td>
          <td width="10%"> <div align="center">至年月</div></td>
          <td width="30%"><div align="center">工作单位</div></td>
          <td width="10%"><div align="center">职务</div></td>
          <td width="10%"><div align="center">月薪</div></td>
          <td width="30%"><div align="center">离职原因</div></td>
        </tr>
        <tr> 
          <td width="10%"><div align="center"><%=rs("d01")%></div></td>
          <td width="10%"><div align="center"><%=rs("d02")%></div></td>
          <td width="30%"><div align="center"><%=rs("d03")%></div></td>
          <td><div align="center"><%=rs("d04")%></div></td>
          <td width="10%"><div align="center"><%=rs("d05")%></div></td>
          <td><%=rs("d06")%></td>
        </tr>
        <tr> 
          <td width="10%"><div align="center"><%=rs("e01")%></div></td>
          <td width="10%"><div align="center"><%=rs("e02")%></div></td>
          <td width="30%"><div align="center"><%=rs("e03")%></div></td>
          <td><div align="center"><%=rs("e04")%></div></td>
          <td width="10%"><div align="center"><%=rs("e05")%></div></td>
          <td><div align="left"><%=rs("e06")%></div></td>
        </tr>
        <tr> 
          <td width="10%"><div align="center"><%=rs("f01")%></div></td>
          <td width="10%"><div align="center"><%=rs("f02")%></div></td>
          <td width="30%"><div align="center"><%=rs("f03")%></div></td>
          <td><div align="center"><%=rs("f04")%></div></td>
          <td width="10%"><div align="center"><%=rs("f05")%></div></td>
          <td><%=rs("f01")%></td>
        </tr>
        <tr> 
          <td width="10%">&nbsp;</td>
          <td width="10%">&nbsp;</td>
          <td width="30%">&nbsp;</td>
          <td>&nbsp;</td>
          <td width="10%">&nbsp;</td>
          <td>&nbsp;</td>
        </tr>
      </table></td>
  </tr>
  <tr bgcolor="#CCCCCC"> 
    <td colspan="5" nowrap><strong>四. 家庭状况</strong></td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td colspan="5" nowrap><table width="100%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolordark="#FFFFFF">
        <tr> 
          <td width="10%"> <div align="center">关系</div></td>
          <td width="10%"> <div align="center">姓名</div></td>
          <td width="30%"><div align="center">工作单位</div></td>
          <td width="10%"><div align="center">职务</div></td>
          <td> <div align="center"></div>
            <div align="center">联系电话</div></td>
        </tr>
        <tr> 
          <td width="10%"><div align="center"><%=rs("h01")%></div></td>
          <td width="10%"><div align="center"><%=rs("h02")%></div></td>
          <td width="30%"><div align="center"><%=rs("h03")%></div></td>
          <td><div align="center"><%=rs("h04")%></div></td>
          <td><div align="center"><%=rs("h05")%></div></td>
        </tr>
        <tr> 
          <td width="10%"><div align="center"><%=rs("i01")%></div></td>
          <td width="10%"><div align="center"><%=rs("i02")%></div></td>
          <td width="30%"><div align="center"><%=rs("i03")%></div></td>
          <td><div align="center"><%=rs("i04")%></div></td>
          <td><div align="center"><%=rs("i05")%></div></td>
        </tr>
        <tr> 
          <td width="10%"><%=rs("k01")%></td>
          <td width="10%"><%=rs("k02")%></td>
          <td width="30%"><%=rs("k03")%></td>
          <td><%=rs("k04")%></td>
          <td><%=rs("k05")%></td>
        </tr>
      </table></td>
  </tr>
  <tr bgcolor="#CCCCCC"> 
    <td colspan="5" nowrap><strong>个人专长或突出业绩：</strong></td>
  </tr>
  <tr> 
    <td colspan="5" nowrap><%=rs("j25")%></td>
  </tr>
</table>
<table width="94%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td align="right"><A href="javascript:window.print();">[打印]</A>&nbsp&nbsp[<a href="javascript:history.back();">返回</a>]</td>
  </tr>
</table>
</body>
</html>
<%
conn.close
set conn=nothing
%>