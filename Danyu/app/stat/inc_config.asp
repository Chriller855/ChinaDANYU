<%
'������Ϣ
countname	= "��վͳ��"	' ͳ��������
yesleft		= 1				' �Ƿ���ʾ���ĵ�����1��ʾ�ǣ�0��ʾ��
yestop		= 0				' �ᵼ��������λ�ã�0��ʾ���ţ�1��ʾ��ҳ����2��ʾ��ҳ��
yesto		= 0				' �Ƿ��ڰ�Ȩ������ʾ����Ȩ��������ʹ��������1��ʾ��ʾ��0��ʾ����ʾ
connpath	= "count.mdb"	' ���ݿ���λ�ã�������д���index.asp�����·��
ipconnpath	= "ip.mdb"		' IP����λ�á�
bakconnpath	= "stats.mdb"	' �󱸿���λ�á�
whatcan		= 4				' �ÿ͵�Ȩ�޵ȼ����йصȼ������뿴 readme �ĵ�
CookieExpires=100			' cookie����ʱ��(��),Ĭ��Ϊ100��
addtime		= 0				' ������ʱ��������ʱ��ͱ���ʱ�����٣���λСʱ
							' ���������ʱ����17:00������ʱ��19:00��������2
							' ������ʱ��17:00������ʱ��15:00��������-2
old_count	= 0				' ��ʹ�ñ�ͳ��ϵͳǰ�ķ�����

							' ע�⣬��ˢ�º�����ͳ�ƶ�������ͳ����Ч��

is_ipcheck	= true			' �Ƿ����÷�ˢ�¹��ܣ�true��ʾ���ã�false��ʾ������
is_online	= false			' �Ƿ�ͳ������������true��ʾ���ã�false��ʾ������

onlythesite1= ""			' ����Ƕ�����λ�ã����ֿհױ�ʾ��������
onlythesite2= ""			' ����Ƕ�����λ�ã����ֿհױ�ʾ��������
'���Ϊ��ʱ��ϵͳ��ִ��ʱ���鱻����ҳ���URL�Ƿ����������ֿ�ͷ��
'�������������ֿ�ͷ����ַһ�ɲ�������
'���� onlythesite1="http://www.01demo.com" �����������ڷ�http://www.01demo.com��ҳ����
'Ƕ����룬��ͳ��������������


'��۲���
FlashWidth	= 130			' FLASHͼ����
FlashHeight	= 58			' FLASHͼ��߶�
mPageSize	= 15			' ÿҳ��¼�����������ҳ��ҳ�棩
mPrecision	= 5				' ��ȷ��С��������λ


'վ����Ϣ
mURL		= "http://www.01demo.com"	' վ������
mName		= "01demo"				' վ������
mNameEn		= "01demo.com"			' վ��Ӣ������
masterName	= "01demo"					' ����Ա��
masterpsw	= "admin"					' ����Ա����
masterEmail	= "zhangyu@01demo.com"			' ����Ա�����ʼ�
SiteBrief	= ":)"	' վ����ܣ�����ʹ��Ӣ�ĵ�˫���źͻ���

'conn
set conn=server.createobject("adodb.connection")
DBPath = Server.MapPath(connpath)
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath


%>