document.observe('dom:loaded', function(){
	news_current();	
})

//导航定位
function news_current(){
//alert('1212')
var wintitle=window.document.title;
if(wintitle.indexOf("AgroSoft 猪场管理软件")!=-1){
	$('menu').down('a',0).id="current_m1";
	}
if(wintitle.indexOf("丹育国际")!=-1){
	$('menu').down('a',1).id="current_m2";
	}
if(wintitle.indexOf("丹育种猪")!=-1){
	$('menu').down('a',2).id="current_m3";
	}
}