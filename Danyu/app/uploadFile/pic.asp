<HTML><HEAD><TITLE>插入图片</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<SCRIPT language=JavaScript event=onclick for=Ok>
window.returnValue = a.value+"\""
if (b.value!="") window.returnValue+=" alt=\""+b.value+"\""
if (c.value!="") window.returnValue+=" style=\""+c.value+"\""
if (e.value!="") window.returnValue+=" align=\""+e.value+"\""
window.returnValue+="border=0\""
//alert(window.returnValue)
window.close();
</SCRIPT>
<link rel="stylesheet" type="text/css" href="../index.css">
</HEAD>
<BODY bgColor=menu>
<FORM name=pic>
  <span style="font-size: 9pt">
  <IFRAME name=ad src="upme_Component.htm" frameBorder=0 width="100%" scrolling=no height=25></IFRAME>
  <BR>
  &nbsp;&nbsp;图片类型：jpg,gif,png,bmp，上传限制：2000K。<BR>
  &nbsp;&nbsp;图片地址：
  <INPUT id=a title="可直接输入网上图片的地址，或由上传程序自动产生图片地址" value="" style="WIDTH: 200px;border: 1px solid #0000FF" size="20">
  <BR>
  &nbsp;&nbsp;提示文字：
  <INPUT id=b title=当鼠标移到图片上时，将会出现该提示文字 style="WIDTH: 200px;border: 1px solid #0000FF" size="20">
  <BR>
  &nbsp;&nbsp;特殊效果：
  <select id=c>
    <option selected>不应用
    <option value=filter:blur(add=1,direction=14,strength=15)>模糊效果
    <option value=filter:gray>黑白照片
    <option value=filter:flipv>颠倒显示
    <option value=filter:fliph>左右反转
    <option value=filter:xray>X 光照片
    <option value=filter:invert>颜色反转</option>
  </select>
  &nbsp;&nbsp;位置： 
  <SELECT id=e>
    <OPTION value="" selected>默认
    <OPTION value=left>居左
    <OPTION value=center>居中
    <OPTION value=right>居右</OPTION>
  </SELECT>
  <BR>
  &nbsp;&nbsp;<FONT 
color=red>位置功能能实现图文环绕，但选择后将不能更改</FONT><BR>
  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
  <INPUT id=Ok title=点击“插入”按钮，在编辑器中插入该图片 style="border: 1px solid #0000FF" type=button value=" 插  入 ">
  </span>
</FORM></BODY></HTML>