
function openScript(url, width, height){
	var Win = window.open(url,"openScript",'width=' + width + ',height=' + height + ',resizable=1,scrollbars=yes,menubar=no,status=no' );
}


function strlength(str)
{
var l=str.length;
var teststr="ͬѧ¼";
var t=l;
if (teststr.length==3){
for(i=0;i<l;i++){
		thischar=str.charCodeAt(i);		
		if(thischar>255){
		t=t+1;}
	}
}
return t;

}


var ie = (document.all)? true:false
if (ie){function ctlent(eventobject){if(event.ctrlKey && window.event.keyCode==13){this.document.frmgbk.submit();}}}

function submitonce(theform){
//if IE 4+ or NS 6+
if (document.all||document.getElementById){
//screen thru every element in the form, and hunt down "submit" and "reset"
for (i=0;i<theform.length;i++){
var tempobj=theform.elements[i]
if(tempobj.type.toLowerCase()=="submit"||tempobj.type.toLowerCase()=="reset")
//disable em
tempobj.disabled=true
}
}
}


function CheckAll(form)  {
  for (var i=0;i<form.elements.length;i++)    {
    var e = form.elements[i];
    if (e.type.toLowerCase()=="checkbox"&&e.name != 'chkall')       e.checked = form.chkall.checked; 
   }
  }
