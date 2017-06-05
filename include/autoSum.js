/* This script and many more are available free online at
The JavaScript Source!! http://javascript.internet.com
Created by: Jim Stiles | www.jdstiles.com */
function startCalc(){
  interval = setInterval("calc()",1);
}

function calc(){
  labour = document.form_gra.txtLabour.value;
  parts = document.form_gra.txtParts.value; 
  document.form_gra.txtGST.value = (labour * 0.1) + (parts * 0.1);

  gst = document.form_gra.txtGST.value; 
  document.form_gra.txtTotalCost.value = (labour * 1) + (parts * 1) + (gst * 1);
  
}

function stopCalc(){
  clearInterval(interval);
}