<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!-- #include file="inc_config.asp" -->
<%

'****************** �������ݶ��� ******************
Set rs = Server.CreateObject("ADODB.Recordset")

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<title><%=countname%>������Ա����̨</title>
<link rel="stylesheet" type="text/css" href="../../index.css">
</head>
<body topmargin=5 rightmargin=0 leftmargin=0 vlink=#000000>
<br>
&nbsp; <font style="FONT-SIZE: 16px" 
     >&nbsp;</font>����Ա����̨<br>
<%if isUpgrade() or (isupgrade12()=false) or isupgrade13()=false then%>
<ul>
  <li style="MARGIN-TOP:	 15px; MARGIN-BOTTOM: 5px; MARGIN-LEFT: 50px"> <A href="upgrade1.asp">���ݿ�ṹ����</a>
      <p>���ԭ����1.0�棬��ȷ���Ѿ��ϴ���������ص��ļ���������ɾ��1.0�����ݿ��ļ��е�IP�⣬������1.1��ǰ�汾���߱��ĵ���Ļ����ֶΣ���1.2�汾���߱��ı��ݱ���ֶμ���ʼͳ�������ֶΡ�
          <%end if%>
      </p>
  <li style="MARGIN-TOP:	 15px; MARGIN-BOTTOM: 5px; MARGIN-LEFT: 50px"> ���ݱ��ݺ�����
      <p>ͳ����ʹ��һ��ʱ���Ժ������ʼ�¼���ݿ���úܴ��ⲻ��ռ���˴�������վ�ռ䣬��ʹͳ����������Ч�ʴ�󽵵ͣ�����Ӧ�ö���������ʼ�¼���ݿ⡣
      <p>Ϊ�˱�֤����������ݲ�������ȫ��ʧ��Ӧ��������ǰ��ʹ�ñ��ݹ��߱���Ҫ��������ݣ����ڵ������߲�������û�б��ݹ��ļ�¼��
      <p><A href="bak_step1.asp" target=_blank>���ݱ���&gt;&gt;&gt;</a> &nbsp;&nbsp;<A href="clear_step1.asp" target=_blank>��������&gt;&gt;&gt;</a>
  <li style="MARGIN-TOP:	 15px; MARGIN-BOTTOM: 5px; MARGIN-LEFT: 50px"> <A href="update.asp?type=go&amp;start=yes" >���µ�IP���ݿ����������������</a>
      <p>��������ռ�ô�����Դ���������м�¼����10000�����Ѳ�Ҫ�����ˡ�ִ�и���ʱ��Ļ���᲻��ˢ�£����ǳ�����ִ�У���Լ10���ӿ��Ը���10000����Ϣ�����¿�����ʱ��ֹ�����ɵ�������̨����ִ�С�
      <p>����������������IP���ݿ�ʱʹ�ã�������б�������
            <%if isupdate() then%>
      </p>
    <li style="MARGIN-TOP:	 15px; MARGIN-BOTTOM: 5px; MARGIN-LEFT: 50px"> <A href="update.asp?type=go">�����ϴ�δ��ɵĸ���</a>
        <p>ϵͳ��鵽����һ�ζ������ݿ�ĸ�����δ��ɣ������������ӿɼ������С�
            <%end if%>
        </p>
    <li style="MARGIN-TOP:	 15px; MARGIN-BOTTOM: 5px; MARGIN-LEFT: 50px"> <A href="logout.asp" target=_top>�˳�����̨</a>
        <p>Ϊ��ֹ���Ĺ���Ա��ݱ��������ã����뿪����ʱ�رձ����ڻ��ߵ��������˳����ӡ�</p>
    </li>
</ul>
</body>
</html>
<%

'****************** �ر����ݶ��� ******************
set rs=nothing
conn.Close 
set conn=nothing

'****************** �Զ��庯�� ********************
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