<!-- #include file="../user/chkpas.asp" -->
<!--#include file="config.asp"-->

<%
'const upload_type=1
if upload_type=0 then
%>
<!--#include file="upload.asp"-->
<%
dim upload,file,formName,SavePath,filename,fileExt
dim upNum
dim EnableUpload
dim Forumupload
dim ranNum
dim uploadfiletype
dim msg,founderr
msg=""
founderr=false
EnableUpload=false
SavePath = SaveUpFilesPath   '����ϴ��ļ���Ŀ¼
if right(SavePath,1)<>"/" then SavePath=SavePath&"/" '��Ŀ¼���(/)

	set upload=new upload_file    '�����ϴ�����
	for each formName in upload.file '�г������ϴ��˵��ļ�
		set file=upload.file(formName)  '����һ���ļ�����
		if file.filesize<100 then
 			msg="����ѡ����Ҫ�ϴ����ļ���"
			founderr=true
		end if
		if file.filesize>(MaxFileSize*1024) then
 			msg="�ļ���С���������ƣ����ֻ���ϴ�" & CStr(MaxFileSize) & "K���ļ���"
			founderr=true
		end if

		fileExt=lcase(file.FileExt)
		Forumupload=split(UpFileType,"|")
		for i=0 to ubound(Forumupload)
			if fileEXT=trim(Forumupload(i)) then
				EnableUpload=true
				exit for
			end if
		next
		if fileEXT="asp" or fileEXT="asa" or fileEXT="aspx" then
			EnableUpload=false
		end if
		if EnableUpload=false then
			msg="�����ļ����Ͳ������ϴ���\n\nֻ�����ϴ��⼸���ļ����ͣ�" & UpFileType
			founderr=true
		end if
		
		strJS="<SCRIPT language=javascript>" & vbcrlf
		if founderr<>true then
			randomize
			ranNum=int(900*rnd)+100
			filename=SavePath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum&"."&fileExt
			file.SaveToFile Server.mappath(FileName)   '�����ļ�
			filename1=year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum&"."&fileExt
		   response.write filename



		    msg="�ϴ��ļ��ɹ���"
 	  		strJS=strJS & "var content;"
 	  '		strJS=strJS & "content=parent.document.pic.a.value;"
 	  				FileType=right(fileExt,3)
			select case FileType
			    case "jpg","gif","png","bmp"
					strJS=strJS &"content=content+'" & filename & "';" & vbcrlf
					
				case "mht"
					strJS=strJS & "content=content+'" & filename & "';" & vbcrlf
					
                case "swf"
					strJS=strJS & "content=content+'" & filename & "';" & vbcrlf
				
	  			end select
	  '			strJS=strJS & "parent.document.pic.a.value=content;" & vbcrlf
	  '		strJS=strJS & "content=parent.document.pic.f.value;"
 	  				FileType=right(fileExt,3)
			select case FileType
			    case "jpg","gif","png","bmp"
					strJS=strJS &"content=content+'" & filename1 & "';" & vbcrlf
					
				case "mht"
					strJS=strJS & "content=content+'" & filename1 & "';" & vbcrlf
					
                case "swf"
					strJS=strJS & "content=content+'" & filename1 & "';" & vbcrlf
				
	  			end select
	  '			strJS=strJS & "parent.document.pic.f.value=content;" & vbcrlf
	
		end if
		strJS=strJS & "alert('" & msg & "');" & vbcrlf
	  	strJS=strJS & "window.location = 'upme_noComponent.htm';" & vbcrlf
	  	strJS=strJS & "</script>"
	  	response.write strJS
		set file=nothing
	next
	set upload=nothing




else
	'Wfolder="/app/uploadFile/upFiles" 
	Wfolder=SaveUpFilesPath
	Set Upload = Server.CreateObject("Persits.Upload.1")
	'Upload.Save "c:\upload"
	NumberOfFiles= Upload.SaveVirtual(Wfolder) 
	
	For Each File in Upload.Files
	'Response.Write File.Name & "= " & File.Path & " (" & File.Size &" bytes)<BR>"
	fileName= Wfolder & "/" & right(File.Path,Len(File.Path)-InStrRev(File.Path,"\",-1,0))
	Next
		strJS="<SCRIPT language=javascript>" & vbcrlf
		   response.write filename
		    msg="�ϴ��ļ��ɹ���"
			strJS=strJS & "content=parent.document.pic.a.value;"
			strJS=strJS &"content=content+'" & filename & "';" & vbcrlf
			strJS=strJS & "parent.document.pic.a.value=content;" & vbcrlf
			'strJS=strJS & "content=parent.document.pic.f.value;"
			'strJS=strJS & "parent.document.pic.f.value=content;" & vbcrlf
		strJS=strJS & "alert('" & msg & "');" & vbcrlf
	  	strJS=strJS & "window.location = 'upme_Component.htm';" & vbcrlf
	  	strJS=strJS & "</script>"
	  	response.write strJS
end if
%>