//document.observe('dom:loaded', function(){
//	news_current();	
//})
window.onload=function(){
news_current();
}

//������λ
function news_current(){
//alert('1212')
var wintitle=window.document.title;
if(wintitle.indexOf("�������")!=-1){
	$('menu').down('a',0).id="current_m1";
	}
if(wintitle.indexOf("����ũҵ")!=-1){
	$('menu').down('a',1).id="current_m2";
	}
}