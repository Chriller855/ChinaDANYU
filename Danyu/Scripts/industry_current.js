document.observe('dom:loaded', function(){
	news_current();	
})

//������λ
function news_current(){
//alert('1212')
var wintitle=window.document.title;
if(wintitle.indexOf("�Ϻ�������г�")!=-1){
	$('menu').down('a',0).id="current_m1";
	}
if(wintitle.indexOf("���ǹɷ�")!=-1){
	$('menu').down('a',1).id="current_m2";
	}
if(wintitle.indexOf("̩���̳�")!=-1){
	$('menu').down('a',2).id="current_m3";
	}
if(wintitle.indexOf("��Ľ�����")!=-1){
	$('menu').down('a',3).id="current_m4";
	}
}