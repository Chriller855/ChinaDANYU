<%
'�ж��û��Ƿ���Ȩ�޷��ʸ�ҳ��ĺ���
'AccessRightStr��userTypeInfo.typeID���ַ�����,��:AccessRightStr="1,2,3"��ʾtypeID=1��2��3���û������Է���
'AccessRightStrΪ�ձ�ʾ����ע���û������Է���
function isAccess(AccessRightStr)
	dim temp,i
	if isNull(AccessRightStr) or AccessRightStr="" then
		isAccess=true
	else
		isAccess=false
		if trim(session("userLevel"))="1" then isAccess=true 'admin�û�ʼ�ն��з���Ȩ
		temp=split(AccessRightStr,",")
		for i=0 to ubound(temp)
			if trim(temp(i))=trim(session("userLevel")) then
				isAccess=true
				exit for
			end if
		next
	end if
end function

if isAccess(iAccessRight)=false then
response.write "<script language='javascript'>alert('���Թ���Ա�˻���¼!');top.document.location.href='../'</script>"

session("loginFlag")=""
session("userlevel")=""
response.end
end if
%>
