/* car brand searching script  */
var a;
function srchBrand(varTxt){
	if(varTxt.length>0){
		var b="ajaxPage3a.asp?sid="+Math.random()+"&qAction=Search&qBrand="+varTxt; 
		a=GetXmlHttpObject(recResult);
		document.getElementById("brandDiv").style.display='';
		document.getElementById("brandDiv").innerHTML="<img src='/images/arrow.gif'/>";
		a.open("GET",b,true);a.send(null);}
	else{
		document.getElementById("brandDiv").style.display='none';
		document.getElementById("brandDiv").innerHTML="";}}

function recResult(){
	if(a.readyState==4||a.readyState=="complete"){
		document.getElementById("brandDiv").innerHTML=a.responseText;}}

/* asynchronous javascript object */
function GetXmlHttpObject(handler){
	var d=null;
	if(navigator.userAgent.indexOf("MSIE")>=0){
		var e="Msxml2.XMLHTTP";
		if(navigator.appVersion.indexOf("MSIE 5.5")>=0){
			e="Microsoft.XMLHTTP";}
		try{
			d=new ActiveXObject(e);
			d.onreadystatechange=handler;
			return d;}
		catch(e){
			alert("Browser Error. Unable to perform AJAX feature");
			return;}}
			
	if(navigator.userAgent.indexOf("Mozilla")>=0 || navigator.userAgent.indexOf("Opera")>=0){
		d=new XMLHttpRequest();
		d.onload=handler;
		d.onerror=handler;
		return d;}}