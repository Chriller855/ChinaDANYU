document.observe('dom:loaded', function(){
	news_current();	
})

//导航定位
function news_current(){
//alert('1212')
var wintitle=window.document.title;
if(wintitle.indexOf("服务理念")!=-1){
	$('menu').down('a',0).id="current_m1";
	}
if(wintitle.indexOf("产品介绍")!=-1){
	$('menu').down('a',1).id="current_m2";
	}
}