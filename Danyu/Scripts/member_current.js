document.observe('dom:loaded', function(){
	news_current();	
})

//导航定位
function news_current(){
//alert('1212')
var wintitle=window.document.title;
if(wintitle.indexOf("化工业")!=-1){
	$('menu').down('a',0).id="current_m1";
	}
if(wintitle.indexOf("房地产业")!=-1){
	$('menu').down('a',1).id="current_m2";
	}
if(wintitle.indexOf("高球产业")!=-1){
	$('menu').down('a',2).id="current_m3";
	}
if(wintitle.indexOf("酒店业")!=-1){
	$('menu').down('a',3).id="current_m4";
	}
}