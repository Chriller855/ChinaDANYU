<!-- www.00ok.net ��Ȩ���� -->
<!-- #include file="iAccessRight.asp" -->
<!-- #include file="../user/chkAdminPas.asp" -->
<!--#include file="conn.asp"-->
<%
	sql="select count(*) from book1 where subid ="&request("subid")
	set rs=conn.execute(sql)
	'response.write rs(0)
		if rs(0)=0 then 
			sql="delete from book1 where id="&request("subid")
			conn.execute(sql)
			rs.close
			set rs=nothing
			conn.close
			set conn=nothing
			response.redirect session("notePadRefererString")
		else
			rs.close
			set rs=nothing
			conn.close
			set conn=nothing
			response.write "<script language=javascript>alert('���и���������£��㲻��ɾ��������');history.go(-1);</script>"
		end if
%>
