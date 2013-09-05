document.observe('dom:loaded', function(){
	news_current();	
})

//导航定位
function news_current(){
//alert('1212')
var wintitle=window.document.title;
if(wintitle.indexOf("丹育种猪繁育计划")!=-1){
	$('menu').down('a',0).id="current_m1";
	}
if(wintitle.indexOf("育种历史")!=-1){
	$('menu').down('a',1).id="current_m2";
	}
if(wintitle.indexOf("育种结构")!=-1){
	$('menu').down('a',2).id="current_m3";
	}
if(wintitle.indexOf("育种目标")!=-1){
	$('menu').down('a',3).id="current_m4";
	}
if(wintitle.indexOf("遗传改良")!=-1){
	$('menu').down('a',4).id="current_m5";
	}
if(wintitle.indexOf("测定结果")!=-1){
	$('menu').down('a',5).id="current_m6";
	}
if(wintitle.indexOf("健康水平")!=-1){
	$('menu').down('a',6).id="current_m7";
	}
if(wintitle.indexOf("屠宰品质")!=-1){
	$('menu').down('a',7).id="current_m8";
	}
}