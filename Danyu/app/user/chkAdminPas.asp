<%
'判断用户是否有权限访问该页面的函数
'AccessRightStr是userTypeInfo.typeID的字符串集,例:AccessRightStr="1,2,3"表示typeID=1、2、3的用户都可以访问
'AccessRightStr为空表示所有注册用户都可以访问
function isAccess(AccessRightStr)
	dim temp,i
	if isNull(AccessRightStr) or AccessRightStr="" then
		isAccess=true
	else
		isAccess=false
		if trim(session("userLevel"))="1" then isAccess=true 'admin用户始终都有访问权
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
response.write "<script language='javascript'>alert('请以管理员账户登录!');top.document.location.href='../'</script>"

session("loginFlag")=""
session("userlevel")=""
response.end
end if
%>
