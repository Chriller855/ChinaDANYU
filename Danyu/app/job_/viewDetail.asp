<!-- www.00ok.net ��Ȩ���� -->
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
    <td align="right"><div align="left">ӦƸְλ��<%=rs("j01")%></div></td>
    <td align="right"><div align="center">ϣ��н�ʴ�����<%=rs("j02")%></div></td>
    <td align="right">�Ǽ�ʱ�䣺<%=rs("j03")%></td>
  </tr>
</table>
<table width="94%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolordark="#FFFFFF">
  <tr bgcolor="#D5D5D5"> 
    <td colspan="5" nowrap> <strong>һ. ��������</strong></td>
  </tr>
  <tr> 
    <td colspan="3" nowrap> ������<%=rs("j04")%></td>
    <td width="36%" nowrap>�Ա�<%=rs("j05")%></td>
    <td width="29%" nowrap>���䣺<%=rs("j06")%></td>
  </tr>
  <tr> 
    <td colspan="3" nowrap>���᣺<%=rs("j07")%></td>
    <td width="36%" nowrap>����״����<%=rs("j08")%></td>
    <td width="29%" nowrap>������Ů��<%=rs("j09")%></td>
  </tr>
  <tr> 
    <td colspan="3" nowrap>������ò��<%=rs("j10")%></td>
    <td width="36%" nowrap>���ѧ����<%=rs("j11")%></td>
    <td width="29%" nowrap>��ߣ�<%=rs("j12")%></td>
  </tr>
  <tr> 
    <td colspan="5" nowrap>����סַ��<%=rs("j13")%></td>
  </tr>
  <tr> 
    <td colspan="3" nowrap>�����ֻ���<%=rs("j14")%></td>
    <td width="36%" nowrap>E-mail��<%="<a href='mailto:"&rs("j15")&"'>"&rs("j15")&"</a>"%></td>
    <td width="29%" nowrap>��ʱ���Կ�ʼ������<%=rs("j16")%></td>
  </tr>
  <tr bgcolor="#F6F6F6"> 
    <td colspan="5" nowrap>���������ϵ��</td>
  </tr>
  <tr bgcolor="#F6F6F6"> 
    <td colspan="2" nowrap>������<%=rs("j17")%></td>
    <td width="13%" nowrap>��ϵ��<%=rs("j18")%></td>
    <td width="36%" nowrap>��ַ��<%=rs("j19")%></td>
    <td width="29%" nowrap>�绰��<%=rs("j20")%></td>
  </tr>
  <tr bgcolor="#CCCCCC"> 
    <td colspan="5" nowrap><strong>��. רҵ�������̶�</strong></td>
  </tr>
  <tr bgcolor="#EFEFEF"> 
    <td width="11%" nowrap><div align="center">������</div></td>
    <td width="11%" nowrap><div align="center">������</div></td>
    <td nowrap><div align="center">ѧ�����</div></td>
    <td width="36%" nowrap><div align="center">��ҵѧУ</div></td>
    <td width="29%" nowrap><div align="center">ѧϰרҵ</div></td>
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
    <td colspan="3" nowrap>���˳�����������֤�飺<%=rs("j21")%></td>
    <td width="36%" nowrap>������������� ���֣�<%=rs("j22")%></td>
    <td width="29%" nowrap>���ճ̶ȣ�<%=rs("j23")%></td>
  </tr>
  <tr> 
    <td colspan="5" nowrap>������������Ӧ�������</td>
  </tr>
  <tr> 
    <td colspan="5" nowrap bgcolor="#CCCCCC"><strong>��. ��������</strong></td>
  </tr>
  <tr> 
    <td colspan="5" nowrap><table width="100%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolordark="#FFFFFF">
        <tr> 
          <td width="10%"> <div align="center">������</div></td>
          <td width="10%"> <div align="center">������</div></td>
          <td width="30%"><div align="center">������λ</div></td>
          <td width="10%"><div align="center">ְ��</div></td>
          <td width="10%"><div align="center">��н</div></td>
          <td width="30%"><div align="center">��ְԭ��</div></td>
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
    <td colspan="5" nowrap><strong>��. ��ͥ״��</strong></td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td colspan="5" nowrap><table width="100%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolordark="#FFFFFF">
        <tr> 
          <td width="10%"> <div align="center">��ϵ</div></td>
          <td width="10%"> <div align="center">����</div></td>
          <td width="30%"><div align="center">������λ</div></td>
          <td width="10%"><div align="center">ְ��</div></td>
          <td> <div align="center"></div>
            <div align="center">��ϵ�绰</div></td>
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
    <td colspan="5" nowrap><strong>����ר����ͻ��ҵ����</strong></td>
  </tr>
  <tr> 
    <td colspan="5" nowrap><%=rs("j25")%></td>
  </tr>
</table>
<table width="94%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td align="right"><A href="javascript:window.print();">[��ӡ]</A>&nbsp&nbsp[<a href="javascript:history.back();">����</a>]</td>
  </tr>
</table>
</body>
</html>
<%
conn.close
set conn=nothing
%>