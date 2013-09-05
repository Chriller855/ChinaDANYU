document.observe('dom:loaded', function(){
	news_current();	
})

//导航定位
function news_current(){
//alert('1212')
var wintitle=window.document.title;
if(wintitle.indexOf("人才理念")!=-1){
	$('menu').down('a',0).id="current_m1";
	}
if(wintitle.indexOf("待聘岗位")!=-1){
	$('menu').down('a',1).id="current_m2";
	}
if(wintitle.indexOf("在线简历")!=-1){
	$('menu').down('a',2).id="current_m3";
	}
}