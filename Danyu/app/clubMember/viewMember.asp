<!-- www.00ok.net ��Ȩ���� -->
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
    <td align="right">ע��ʱ�䣺<%=rs("dateAndTime")%></td>
  </tr>
</table>
<table width="94%" border="1" cellspacing="0" cellpadding="0" bordercolordark="#FFFFFF">
  <tr> 
    <td width="242" nowrap> *��������/Name��</td>
    <td width="305"><%=rs("memberName")%></td>
  </tr>
  <tr> 
    <td nowrap> *�Ա�/Gender��</td>
    <td><%=rs("sexid")%></td>
  </tr>
  <tr> 
    <td nowrap> *��������/Date of birth��</td>
    <td> <%=rs("birthday")%></td>
  </tr>
  <tr> 
    <td nowrap>����/Nationality��</td>
    <td> <%=rs("Nationality")%></td>
  </tr>
  <tr> 
    <td nowrap>����/Place of Birth��</td>
    <td> <%=rs("placeOfBirth")%></td>
  </tr>
  <tr> 
    <td nowrap> *���֤����/ID/Passport No��</td>
    <td> <%=rs("idNum")%></td>
  </tr>
  <tr> 
    <td nowrap> *ͨѶ��ַ/Mail address��</td>
    <td> <%=rs("address")%></td>
  </tr>
  <tr> 
    <td nowrap> *��������Postal code��</td>
    <td> <%=rs("postCode")%></td>
  </tr>
  <tr> 
    <td nowrap> ��ϵ�绰/Phone No��</td>
    <td> סլ/ Home�� <%=rs("telHome")%><br>
      �ֻ�/Mobil�� <%=rs("Mobile")%></td>
  </tr>
  <tr> 
    <td nowrap> �����ʼ�/E-mail address��</td>
    <td> <%=rs("email")%></td>
  </tr>
  <tr> 
    <td nowrap>��ְ��˾/Place of Employment��</td>
    <td> <%=rs("PlaceOfEmployment")%></td>
  </tr>
  <tr> 
    <td nowrap> ְҵ/Occupation��</td>
    <td> <%=rs("occupation")%></td>
  </tr>
  <tr> 
    <td nowrap> �����̶�/Education Degree��</td>
    <td> <%=rs("education")%> </td>
  </tr>
  <tr> 
    <td nowrap>���򹺷��ۿ�/Intended Price��</td>
    <td> <%=rs("IntendedPrice")%> </td>
  </tr>
  <tr> 
    <td nowrap>���򹺷���׼/Intended Requirements��</td>
    <td> <%=rs("IntendedRequirements")%></td>
  </tr>
  <tr> 
    <td nowrap>���򹺷���;/Intended Usage��</td>
    <td> <%=rs("IntendedUsage")%></td>
  </tr>
  <tr> 
    <td nowrap>�ⶨ��������/Intended District��</td>
    <td><%=rs("IntendedDistrict")%></td>
  </tr>
  <tr> 
    <td nowrap> ��Ȥ����/Hobbies��</td>
    <td><%=rs("interest")%></td>
  </tr>
  <tr> 
    <td nowrap>������ĻἮ���/Member Type��</td>
    <td> <%=rs("MemberType")%></td>
  </tr>
  <tr> 
    <td nowrap>������ҵ����Ա��<br>
      ����ǰ�����õĻ����õأ�</td>
    <td>¥������/CRLand Real Estate You Purchased�� <%=rs("BuildingName")%><br>
      ¥����/Building Number�� <%=rs("BuildingNumber")%></td>
  </tr>
</table>
<table width="94%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td align="right">[<a href="javascript:history.back();">����</a>]</td>
  </tr>
</table>
</body>
</html>
<%
conn.close
set conn=nothing
%>