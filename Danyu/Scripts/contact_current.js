document.observe('dom:loaded', function(){
	news_current();	
})

//导航定位
function news_current(){
//alert('1212')
var wintitle=window.document.title;
if(wintitle.indexOf("联系方式")!=-1){
	$('menu').down('a',0).id="current_m1";
	}
if(wintitle.indexOf("在线咨询")!=-1){
	$('menu').down('a',1).id="current_m2";
	}
}