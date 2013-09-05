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
SavePath = SaveUpFilesPath   '存放上传文件的目录
if right(SavePath,1)<>"/" then SavePath=SavePath&"/" '在目录后加(/)

	set upload=new upload_file    '建立上传对象
	for each formName in upload.file '列出所有上传了的文件
		set file=upload.file(formName)  '生成一个文件对象
		if file.filesize<100 then
 			msg="请先选择你要上传的文件！"
			founderr=true
		end if
		if file.filesize>(MaxFileSize*1024) then
 			msg="文件大小超过了限制，最大只能上传" & CStr(MaxFileSize) & "K的文件！"
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
			msg="这种文件类型不允许上传！\n\n只允许上传这几种文件类型：" & UpFileType
			founderr=true
		end if
		
		strJS="<SCRIPT language=javascript>" & vbcrlf
		if founderr<>true then
			randomize
			ranNum=int(900*rnd)+100
			filename=SavePath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum&"."&fileExt
			file.SaveToFile Server.mappath(FileName)   '保存文件
			filename1=year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum&"."&fileExt
		   response.write filename



		    msg="上传文件成功！"
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
		    msg="上传文件成功！"
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