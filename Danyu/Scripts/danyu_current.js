document.observe('dom:loaded', function(){
	news_current();	
})

//������λ
function news_current(){
//alert('1212')
var wintitle=window.document.title;
if(wintitle.indexOf("�����������ƻ�")!=-1){
	$('menu').down('a',0).id="current_m1";
	}
if(wintitle.indexOf("������ʷ")!=-1){
	$('menu').down('a',1).id="current_m2";
	}
if(wintitle.indexOf("���ֽṹ")!=-1){
	$('menu').down('a',2).id="current_m3";
	}
if(wintitle.indexOf("����Ŀ��")!=-1){
	$('menu').down('a',3).id="current_m4";
	}
if(wintitle.indexOf("�Ŵ�����")!=-1){
	$('menu').down('a',4).id="current_m5";
	}
if(wintitle.indexOf("�ⶨ���")!=-1){
	$('menu').down('a',5).id="current_m6";
	}
if(wintitle.indexOf("����ˮƽ")!=-1){
	$('menu').down('a',6).id="current_m7";
	}
if(wintitle.indexOf("����Ʒ��")!=-1){
	$('menu').down('a',7).id="current_m8";
	}
}