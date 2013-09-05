<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>图片上传</title>
<link rel="stylesheet" type="text/css" href="../index.css">
</head>
<body leftmargin="0" topmargin="0">
<form name="form" method="post" action="insertPicAct.asp?obj=<%=request("obj")%>&frmName=<%=request("frmName")%>&picFlag=<%=request("picFlag")%>" enctype="multipart/form-data" >
  <span style="font-size: 9pt">
<input type="file" name="file1" size=20 style="border: 1px solid #888888">
<input type="submit" name="Submit" value="上传" style="border: 1px solid #888888" >
</span>
</form>
</body>
</html>