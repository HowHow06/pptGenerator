
function transs(){
	var cc=document.form2.code3.value; // the field to convert
	var str='';
	var s;
	for(var i=0;i<cc.length;i++){
	//alert(cc.charAt(i)+" = "+cc.charCodeAt(i));
		if(pydis.indexOf(cc.charAt(i))!=-1&&cc.charCodeAt(i)>200){
			s=1;
			while(pydis.charAt(pydis.indexOf(cc.charAt(i))+s)!=","){
				str+=pydis.charAt(pydis.indexOf(cc.charAt(i))+s);
				s++;
			}
			str+=" ";
		}
		else{
			str+=cc.charAt(i);
		}
	}
	return str;
}