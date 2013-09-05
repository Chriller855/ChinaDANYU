<!-- www.00ok.net 版权所有 -->
<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!--#include file="conn.asp"-->
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="clubMember.css" rel="stylesheet" type="text/css">
</head>
<body>
<%
set rs = createObject("adodb.recordset")
sql="select * from clubMember where id="&request("id")
rs.open sql,conn,1,1
%>
<table width="94%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="right">注册时间：<%=rs("dateAndTime")%></td>
  </tr>
</table>
<table width="94%" border="1" cellspacing="0" cellpadding="0" bordercolordark="#FFFFFF">
  <tr> 
    <td width="242" nowrap> *中文姓名/Name：</td>
    <td width="305"><%=rs("memberName")%></td>
  </tr>
  <tr> 
    <td nowrap> *性别/Gender：</td>
    <td><%=rs("sexid")%></td>
  </tr>
  <tr> 
    <td nowrap> *出生日期/Date of birth：</td>
    <td> <%=rs("birthday")%></td>
  </tr>
  <tr> 
    <td nowrap>国籍/Nationality：</td>
    <td> <%=rs("Nationality")%></td>
  </tr>
  <tr> 
    <td nowrap>籍贯/Place of Birth：</td>
    <td> <%=rs("placeOfBirth")%></td>
  </tr>
  <tr> 
    <td nowrap> *身份证号码/ID/Passport No：</td>
    <td> <%=rs("idNum")%></td>
  </tr>
  <tr> 
    <td nowrap> *通讯地址/Mail address：</td>
    <td> <%=rs("address")%></td>
  </tr>
  <tr> 
    <td nowrap> *邮政编码Postal code：</td>
    <td> <%=rs("postCode")%></td>
  </tr>
  <tr> 
    <td nowrap> 联系电话/Phone No：</td>
    <td> 住宅/ Home： <%=rs("telHome")%><br>
      手机/Mobil： <%=rs("Mobile")%></td>
  </tr>
  <tr> 
    <td nowrap> 电子邮件/E-mail address：</td>
    <td> <%=rs("email")%></td>
  </tr>
  <tr> 
    <td nowrap>就职公司/Place of Employment：</td>
    <td> <%=rs("PlaceOfEmployment")%></td>
  </tr>
  <tr> 
    <td nowrap> 职业/Occupation：</td>
    <td> <%=rs("occupation")%></td>
  </tr>
  <tr> 
    <td nowrap> 教育程度/Education Degree：</td>
    <td> <%=rs("education")%> </td>
  </tr>
  <tr> 
    <td nowrap>意向购房价款/Intended Price：</td>
    <td> <%=rs("IntendedPrice")%> </td>
  </tr>
  <tr> 
    <td nowrap>意向购房标准/Intended Requirements：</td>
    <td> <%=rs("IntendedRequirements")%></td>
  </tr>
  <tr> 
    <td nowrap>意向购房用途/Intended Usage：</td>
    <td> <%=rs("IntendedUsage")%></td>
  </tr>
  <tr> 
    <td nowrap>拟定购房区域/Intended District：</td>
    <td><%=rs("IntendedDistrict")%></td>
  </tr>
  <tr> 
    <td nowrap> 兴趣爱好/Hobbies：</td>
    <td><%=rs("interest")%></td>
  </tr>
  <tr> 
    <td nowrap>您申请的会籍类别/Member Type：</td>
    <td> <%=rs("MemberType")%></td>
  </tr>
  <tr> 
    <td nowrap>如申请业主会员，<br>
      您以前所购置的华润置地：</td>
    <td>楼盘名称/CRLand Real Estate You Purchased： <%=rs("BuildingName")%><br>
      楼座号/Building Number： <%=rs("BuildingNumber")%></td>
  </tr>
</table>
<table width="94%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td align="right">[<a href="javascript:history.back();">返回</a>]</td>
  </tr>
</table>
</body>
</html>
<%
conn.close
set conn=nothing
%>