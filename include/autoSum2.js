/* This script and many more are available free online at
The JavaScript Source!! http://javascript.internet.com
Created by: Jim Stiles | www.jdstiles.com */
function startCalc(){
  interval = setInterval("calc()",1);
}

function calc(){
  height = document.form_calculator.txtHeight.value;
  width = document.form_calculator.txtWidth.value;
  depth = document.form_calculator.txtDepth.value; 
   
  document.form_calculator.txtTotal.value = (height * width * depth) / 1000000;
  
}

function stopCalc(){
  clearInterval(interval);
}