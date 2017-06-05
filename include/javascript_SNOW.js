<!--
function close_window() {
  if (confirm("Close Window?")) {
    close();
  }
}


function GetDay(nDay)
{
	var Days = new Array("Sunday","Monday","Tuesday","Wednesday",
	                     "Thursday","Friday","Saturday");
	return Days[nDay]
}

function GetMonth(nMonth)
{
	var Months = new Array("January","February","March","April","May","June",
	                       "July","August","September","October","November","December");
	return Months[nMonth]
}

function DateString()
{
	var Today = new Date();
	var suffix = "th";
	switch (Today.getDate())
	{
		case 1:
		case 21:
		case 31:
			suffix = "st"; break;
		case 2:
		case 22:
			suffix = "nd"; break;
		case 3:
		case 23:
			suffix = "rd"; break;
	};

	var strDate = GetDay(Today.getDay()) + " " + Today.getDate();
	strDate += suffix + " " + GetMonth(Today.getMonth()) + ", " + Today.getYear();
	return strDate
}

/*******************************************************************************************
MENU
*******************************************************************************************/

var xcNode = [];

function xcSet(m, c) 
{
	if (document.getElementById && document.createElement) 
	{
		m = document.getElementById(m).getElementsByTagName('ul');
		var d, p, x, h, i, j;

		for (i = 0; i < m.length; i++) 
		{
			if (d = m[i].getAttribute('id')) 
			{
				xcCtrl(d, c, 'x', '[+]', 'Show', m[i].getAttribute('title')+' (expand menu)');
				x = xcCtrl(d, c, 'c', '[-]', 'Hide', m[i].getAttribute('title')+' (collapse menu)');

				p = m[i].parentNode;
				
				if (h = !p.className) 
				{
					j = 2;
					while ((h = !(d == arguments[j])) && (j++ < arguments.length));
						if (h) 
						{
							m[i].style.display = 'none';
							x = xcNode[d+'x'];
						}
				}

				p.className = c;
				p.insertBefore(x, p.firstChild);
			}
		}
	}
}


function xcShow(m) 
{
	xcXC(m, 'block', m+'c', m+'x');
}



function xcHide(m) 
{
	xcXC(m, 'none', m+'x', m+'c');
}



function xcXC(e, d, s, h) 
{
	e = document.getElementById(e);
	e.style.display = d;
	e.parentNode.replaceChild(xcNode[s], xcNode[h]);
	xcNode[s].firstChild.focus();
}



function xcCtrl(m, c, s, v, f, t) 
{
	var a = document.createElement('a');
	a.setAttribute('href', 'javascript:xc'+f+'(\''+m+'\');');
	a.setAttribute('title', t);
	a.appendChild(document.createTextNode(v));

	var d = document.createElement('div');
	d.className = c+s;
	d.appendChild(a);

	return xcNode[m+s] = d;
}

/*******************************************************************************************
ROLLOVER
*******************************************************************************************/

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_showHideLayers() { //v6.0
  var i,p,v,obj,args=MM_showHideLayers.arguments;

  for (i=0; i<(args.length-2); i+=3) 
	if ((obj=MM_findObj(args[i]))!=null) 
	{ 
		v=args[i+2];
		if (obj.style) 
		{ 
			obj=obj.style; 
			v=(v=='show')?'inline':(v=='hide')?'none':v;
		}
		obj.display=v; 
	}
}

function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}

/*******************************************************************************************
PRINT FRIENDLY
*******************************************************************************************/

function PrintThisPage() 
{ 
   var sOption="toolbar=yes,location=no,directories=yes,menubar=yes,"; 
       sOption+="scrollbars=yes,width=750,height=600,left=100,top=25"; 

   var sWinHTML = document.getElementById('contentstart').innerHTML; 
   
   var winprint=window.open("","",sOption); 
       winprint.document.open(); 
       winprint.document.write('<html><LINK href=../include/stylesheet_merged.css rel=Stylesheet><body class=printfriendly>'); 
       winprint.document.write(sWinHTML);          
       winprint.document.write('</body></html>'); 
       winprint.document.close(); 
       winprint.focus(); 
}

/*******************************************************************************************
EMAIL THIS PAGE 
*******************************************************************************************/

function mailpage()
{
mail_str = "mailto:?subject=Check out the " + document.title;
mail_str += "&body=I thought you might be interested in the " + document.title;
mail_str += ". You can view it at, " + location.href.toString();
location.href = mail_str;
}

/*******************************************************************************************
Gallery script
*******************************************************************************************/

function LoadGallery(pictureName,imageFile)
{
  if (document.all)
  {
    document.getElementById(pictureName).style.filter="blendTrans(duration=1)";
    document.getElementById(pictureName).filters.blendTrans.Apply();
  }
  document.getElementById(pictureName).src = imageFile;
  if (document.all)
  {
    document.getElementById(pictureName).filters.blendTrans.Play();
  }

}

function validateEmpty(fld) {
    var error = "";
 
    if (fld.value.length == 0) {
        fld.style.background = 'Yellow'; 
        error = "The required field has not been filled in.\n"
    } else {
        fld.style.background = 'White';
    }
    return error;  
}

function validateFormOnSubmit(theForm) {
	var reason = "";
	
	  reason += validateEmpty(theForm.company);
	  reason += validateEmpty(theForm.account_manager);
	  reason += validateEmpty(theForm.sales);
	  reason += validateEmpty(theForm.alliance_partner);
	  		  
	  if (reason != "") {
		alert("Some fields need correction:\n" + reason);
		return false;
	  }
	
	  return true;
}


var xmlHttp

function showHint(str)
{
	if (str.length==0)
	  { 
	  document.getElementById("txtHint").innerHTML="";
	  return;
	  }
	xmlHttp=GetXmlHttpObject();
	if (xmlHttp==null)
	  {
	  alert ("Your browser does not support AJAX!");
	  return;
	  } 
	var url="gethint.asp";
	url=url+"?q="+str;
	url=url+"&sid="+Math.random();
	xmlHttp.onreadystatechange=stateChanged;
	xmlHttp.open("GET",url,true);
	xmlHttp.send(null);
} 

function stateChanged() 
{ 
		if (xmlHttp.readyState==4)
	{ 
	document.getElementById("txtHint").innerHTML=xmlHttp.responseText;
	}
}

function GetXmlHttpObject()
{
	var xmlHttp=null;
	try
	  {
	  // Firefox, Opera 8.0+, Safari
	  xmlHttp=new XMLHttpRequest();
	  }
	catch (e)
	  {
	  // Internet Explorer
	  try
		{
		xmlHttp=new ActiveXObject("Msxml2.XMLHTTP");
		}
	  catch (e)
		{
		xmlHttp=new ActiveXObject("Microsoft.XMLHTTP");
		}
	  }
	return xmlHttp;
}


/**
 * DHTML date validation script for dd/mm/yyyy. Courtesy of SmartWebby.com (http://www.smartwebby.com/dhtml/)
 */
// Declaring valid date character, minimum year and maximum year
var dtCh= "/";
var minYear=1900;
var maxYear=2100;

function isInteger(s){
	var i;
    for (i = 0; i < s.length; i++){   
        // Check that current character is number.
        var c = s.charAt(i);
        if (((c < "0") || (c > "9"))) return false;
    }
    // All characters are numbers.
    return true;
}

function stripCharsInBag(s, bag){
	var i;
    var returnString = "";
    // Search through string's characters one by one.
    // If character is not in bag, append to returnString.
    for (i = 0; i < s.length; i++){   
        var c = s.charAt(i);
        if (bag.indexOf(c) == -1) returnString += c;
    }
    return returnString;
}

function daysInFebruary (year){
	// February has 29 days in any year evenly divisible by four,
    // EXCEPT for centurial years which are not also divisible by 400.
    return (((year % 4 == 0) && ( (!(year % 100 == 0)) || (year % 400 == 0))) ? 29 : 28 );
}
function DaysArray(n) {
	for (var i = 1; i <= n; i++) {
		this[i] = 31
		if (i==4 || i==6 || i==9 || i==11) {this[i] = 30}
		if (i==2) {this[i] = 29}
   } 
   return this
}

function isDate(dtStr){
	var daysInMonth = DaysArray(12)
	var pos1=dtStr.indexOf(dtCh)
	var pos2=dtStr.indexOf(dtCh,pos1+1)
	var strDay=dtStr.substring(0,pos1)
	var strMonth=dtStr.substring(pos1+1,pos2)
	var strYear=dtStr.substring(pos2+1)
	strYr=strYear
	if (strDay.charAt(0)=="0" && strDay.length>1) strDay=strDay.substring(1)
	if (strMonth.charAt(0)=="0" && strMonth.length>1) strMonth=strMonth.substring(1)
	for (var i = 1; i <= 3; i++) {
		if (strYr.charAt(0)=="0" && strYr.length>1) strYr=strYr.substring(1)
	}
	month=parseInt(strMonth)
	day=parseInt(strDay)
	year=parseInt(strYr)
	if (pos1==-1 || pos2==-1){
		//alert("The date format should be : dd/mm/yyyy")
		return false
	}
	if (strMonth.length<1 || month<1 || month>12){
		alert("Please enter a valid month")
		return false
	}
	if (strDay.length<1 || day<1 || day>31 || (month==2 && day>daysInFebruary(year)) || day > daysInMonth[month]){
		alert("Please enter a valid day")
		return false
	}
	if (strYear.length != 4 || year==0 || year<minYear || year>maxYear){
		alert("Please enter a valid 4 digit year between "+minYear+" and "+maxYear)
		return false
	}
	if (dtStr.indexOf(dtCh,pos2+1)!=-1 || isInteger(stripCharsInBag(dtStr, dtCh))==false){
		alert("Please enter a valid date")
		return false
	}
return true
}

function IsNumericNoNull(strString)
       //  check for valid numeric strings	
       {
       var strValidChars = "0123456789.-";
       var strChar;
       var blnResult = true;

       if (strString.length == 0) return false;

       //  test strString consists of valid characters listed above
       for (i = 0; i < strString.length && blnResult == true; i++)
          {
          strChar = strString.charAt(i);
          if (strValidChars.indexOf(strChar) == -1)
             {
             blnResult = false;
             }
          }
       return blnResult;
       }
	   
function IsNumeric(strString) //  check for valid numeric strings	
{
	if(!/\D/.test(strString)) return true;//IF NUMBER
	else if(/^\d+\.\d+$/.test(strString)) return true;//IF A DECIMAL NUMBER HAVING AN INTEGER ON EITHER SIDE OF THE DOT(.)
	else return false;
}	   

function confirmSubmit(){

	var msg = '';
	msg = msg + '---------------------------------------------------------------\n';
	msg = msg + 'Are you sure you want to submit this timesheet?\n';
	msg = msg + '\nPress cancel if you do not wish to submit.\n';
	msg = msg + '---------------------------------------------------------------\n';
	return confirm(msg);
}

function openWindow(anchor, options) {
 
	var args = '';
 
	if (typeof(options) == 'undefined') { var options = new Object(); }
	if (typeof(options.name) == 'undefined') { options.name = 'win' + Math.round(Math.random()*100000); }
 
	if (typeof(options.height) != 'undefined' && typeof(options.fullscreen) == 'undefined') {
		args += "height=" + options.height + ",";
	}
 
	if (typeof(options.width) != 'undefined' && typeof(options.fullscreen) == 'undefined') {
		args += "width=" + options.width + ",";
	}
 
	if (typeof(options.fullscreen) != 'undefined') {
		args += "width=" + screen.availWidth + ",";
		args += "height=" + screen.availHeight + ",";
	}
 
	if (typeof(options.center) == 'undefined') {
		options.x = 0;
		options.y = 0;
		args += "screenx=" + options.x + ",";
		args += "screeny=" + options.y + ",";
		args += "left=" + options.x + ",";
		args += "top=" + options.y + ",";
	}
 
	if (typeof(options.center) != 'undefined' && typeof(options.fullscreen) == 'undefined') {
		options.y=Math.floor((screen.availHeight-(options.height || screen.height))/2)-(screen.height-screen.availHeight);
		options.x=Math.floor((screen.availWidth-(options.width || screen.width))/2)-(screen.width-screen.availWidth);
		args += "screenx=" + options.x + ",";
		args += "screeny=" + options.y + ",";
		args += "left=" + options.x + ",";
		args += "top=" + options.y + ",";
	}
 
	if (typeof(options.scrollbars) != 'undefined') { args += "scrollbars=1,"; }
	if (typeof(options.menubar) != 'undefined') { args += "menubar=1,"; }
	if (typeof(options.locationbar) != 'undefined') { args += "location=1,"; }
	if (typeof(options.resizable) != 'undefined') { args += "resizable=1,"; }
 
	var win = window.open(anchor, options.name, args);
	return false;
 
}

/** @license
 * DHTML Snowstorm! JavaScript-based Snow for web pages
 * --------------------------------------------------------
 * Version 1.42.20111120 (Previous rev: 1.41.20101113)
 * Copyright (c) 2007, Scott Schiller. All rights reserved.
 * Code provided under the BSD License:
 * http://schillmania.com/projects/snowstorm/license.txt
 */

/*global window, document, navigator, clearInterval, setInterval */
/*jslint white: false, onevar: true, plusplus: false, undef: true, nomen: true, eqeqeq: true, bitwise: true, regexp: true, newcap: true, immed: true */

var snowStorm = (function(window, document) {

  // --- common properties ---

  this.flakesMax = 128;           // Limit total amount of snow made (falling + sticking)
  this.flakesMaxActive = 64;      // Limit amount of snow falling at once (less = lower CPU use)
  this.animationInterval = 33;    // Theoretical "miliseconds per frame" measurement. 20 = fast + smooth, but high CPU use. 50 = more conservative, but slower
  this.excludeMobile = true;      // Snow is likely to be bad news for mobile phones' CPUs (and batteries.) By default, be nice.
  this.flakeBottom = null;        // Integer for Y axis snow limit, 0 or null for "full-screen" snow effect
  this.followMouse = true;        // Snow movement can respond to the user's mouse
  this.snowColor = '#fff';        // Don't eat (or use?) yellow snow.
  this.snowCharacter = '&bull;';  // &bull; = bullet, &middot; is square on some systems etc.
  this.snowStick = true;          // Whether or not snow should "stick" at the bottom. When off, will never collect.
  this.targetElement = null;      // element which snow will be appended to (null = document.body) - can be an element ID eg. 'myDiv', or a DOM node reference
  this.useMeltEffect = true;      // When recycling fallen snow (or rarely, when falling), have it "melt" and fade out if browser supports it
  this.useTwinkleEffect = false;  // Allow snow to randomly "flicker" in and out of view while falling
  this.usePositionFixed = false;  // true = snow does not shift vertically when scrolling. May increase CPU load, disabled by default - if enabled, used only where supported

  // --- less-used bits ---

  this.freezeOnBlur = true;       // Only snow when the window is in focus (foreground.) Saves CPU.
  this.flakeLeftOffset = 0;       // Left margin/gutter space on edge of container (eg. browser window.) Bump up these values if seeing horizontal scrollbars.
  this.flakeRightOffset = 0;      // Right margin/gutter space on edge of container
  this.flakeWidth = 8;            // Max pixel width reserved for snow element
  this.flakeHeight = 8;           // Max pixel height reserved for snow element
  this.vMaxX = 5;                 // Maximum X velocity range for snow
  this.vMaxY = 4;                 // Maximum Y velocity range for snow
  this.zIndex = 0;                // CSS stacking order applied to each snowflake

  // --- End of user section ---

  var s = this, storm = this, i,
  // UA sniffing and backCompat rendering mode checks for fixed position, etc.
  isIE = navigator.userAgent.match(/msie/i),
  isIE6 = navigator.userAgent.match(/msie 6/i),
  isWin98 = navigator.appVersion.match(/windows 98/i),
  isMobile = navigator.userAgent.match(/mobile/i),
  isBackCompatIE = (isIE && document.compatMode === 'BackCompat'),
  noFixed = (isMobile || isBackCompatIE || isIE6),
  screenX = null, screenX2 = null, screenY = null, scrollY = null, vRndX = null, vRndY = null,
  windOffset = 1,
  windMultiplier = 2,
  flakeTypes = 6,
  fixedForEverything = false,
  opacitySupported = (function(){
    try {
      document.createElement('div').style.opacity = '0.5';
    } catch(e) {
      return false;
    }
    return true;
  }()),
  didInit = false,
  docFrag = document.createDocumentFragment();

  this.timers = [];
  this.flakes = [];
  this.disabled = false;
  this.active = false;
  this.meltFrameCount = 20;
  this.meltFrames = [];

  this.events = (function() {

    var old = (!window.addEventListener && window.attachEvent), slice = Array.prototype.slice,
    evt = {
      add: (old?'attachEvent':'addEventListener'),
      remove: (old?'detachEvent':'removeEventListener')
    };

    function getArgs(oArgs) {
      var args = slice.call(oArgs), len = args.length;
      if (old) {
        args[1] = 'on' + args[1]; // prefix
        if (len > 3) {
          args.pop(); // no capture
        }
      } else if (len === 3) {
        args.push(false);
      }
      return args;
    }

    function apply(args, sType) {
      var element = args.shift(),
          method = [evt[sType]];
      if (old) {
        element[method](args[0], args[1]);
      } else {
        element[method].apply(element, args);
      }
    }

    function addEvent() {
      apply(getArgs(arguments), 'add');
    }

    function removeEvent() {
      apply(getArgs(arguments), 'remove');
    }

    return {
      add: addEvent,
      remove: removeEvent
    };

  }());

  function rnd(n,min) {
    if (isNaN(min)) {
      min = 0;
    }
    return (Math.random()*n)+min;
  }

  function plusMinus(n) {
    return (parseInt(rnd(2),10)===1?n*-1:n);
  }

  this.randomizeWind = function() {
    var i;
    vRndX = plusMinus(rnd(s.vMaxX,0.2));
    vRndY = rnd(s.vMaxY,0.2);
    if (this.flakes) {
      for (i=0; i<this.flakes.length; i++) {
        if (this.flakes[i].active) {
          this.flakes[i].setVelocities();
        }
      }
    }
  };

  this.scrollHandler = function() {
    var i;
    // "attach" snowflakes to bottom of window if no absolute bottom value was given
    scrollY = (s.flakeBottom?0:parseInt(window.scrollY||document.documentElement.scrollTop||document.body.scrollTop,10));
    if (isNaN(scrollY)) {
      scrollY = 0; // Netscape 6 scroll fix
    }
    if (!fixedForEverything && !s.flakeBottom && s.flakes) {
      for (i=s.flakes.length; i--;) {
        if (s.flakes[i].active === 0) {
          s.flakes[i].stick();
        }
      }
    }
  };

  this.resizeHandler = function() {
    if (window.innerWidth || window.innerHeight) {
      screenX = window.innerWidth-(!isIE?16:16)-s.flakeRightOffset;
      screenY = (s.flakeBottom?s.flakeBottom:window.innerHeight);
    } else {
      screenX = (document.documentElement.clientWidth||document.body.clientWidth||document.body.scrollWidth)-(!isIE?8:0)-s.flakeRightOffset;
      screenY = s.flakeBottom?s.flakeBottom:(document.documentElement.clientHeight||document.body.clientHeight||document.body.scrollHeight);
    }
    screenX2 = parseInt(screenX/2,10);
  };

  this.resizeHandlerAlt = function() {
    screenX = s.targetElement.offsetLeft+s.targetElement.offsetWidth-s.flakeRightOffset;
    screenY = s.flakeBottom?s.flakeBottom:s.targetElement.offsetTop+s.targetElement.offsetHeight;
    screenX2 = parseInt(screenX/2,10);
  };

  this.freeze = function() {
    // pause animation
    var i;
    if (!s.disabled) {
      s.disabled = 1;
    } else {
      return false;
    }
    for (i=s.timers.length; i--;) {
      clearInterval(s.timers[i]);
    }
  };

  this.resume = function() {
    if (s.disabled) {
       s.disabled = 0;
    } else {
      return false;
    }
    s.timerInit();
  };

  this.toggleSnow = function() {
    if (!s.flakes.length) {
      // first run
      s.start();
    } else {
      s.active = !s.active;
      if (s.active) {
        s.show();
        s.resume();
      } else {
        s.stop();
        s.freeze();
      }
    }
  };

  this.stop = function() {
    var i;
    this.freeze();
    for (i=this.flakes.length; i--;) {
      this.flakes[i].o.style.display = 'none';
    }
    s.events.remove(window,'scroll',s.scrollHandler);
    s.events.remove(window,'resize',s.resizeHandler);
    if (s.freezeOnBlur) {
      if (isIE) {
        s.events.remove(document,'focusout',s.freeze);
        s.events.remove(document,'focusin',s.resume);
      } else {
        s.events.remove(window,'blur',s.freeze);
        s.events.remove(window,'focus',s.resume);
      }
    }
  };

  this.show = function() {
    var i;
    for (i=this.flakes.length; i--;) {
      this.flakes[i].o.style.display = 'block';
    }
  };

  this.SnowFlake = function(parent,type,x,y) {
    var s = this, storm = parent;
    this.type = type;
    this.x = x||parseInt(rnd(screenX-20),10);
    this.y = (!isNaN(y)?y:-rnd(screenY)-12);
    this.vX = null;
    this.vY = null;
    this.vAmpTypes = [1,1.2,1.4,1.6,1.8]; // "amplification" for vX/vY (based on flake size/type)
    this.vAmp = this.vAmpTypes[this.type];
    this.melting = false;
    this.meltFrameCount = storm.meltFrameCount;
    this.meltFrames = storm.meltFrames;
    this.meltFrame = 0;
    this.twinkleFrame = 0;
    this.active = 1;
    this.fontSize = (10+(this.type/5)*10);
    this.o = document.createElement('div');
    this.o.innerHTML = storm.snowCharacter;
    this.o.style.color = storm.snowColor;
    this.o.style.position = (fixedForEverything?'fixed':'absolute');
    this.o.style.width = storm.flakeWidth+'px';
    this.o.style.height = storm.flakeHeight+'px';
    this.o.style.fontFamily = 'arial,verdana';
    this.o.style.overflow = 'hidden';
    this.o.style.fontWeight = 'normal';
    this.o.style.zIndex = storm.zIndex;
    docFrag.appendChild(this.o);

    this.refresh = function() {
      if (isNaN(s.x) || isNaN(s.y)) {
        // safety check
        return false;
      }
      s.o.style.left = s.x+'px';
      s.o.style.top = s.y+'px';
    };

    this.stick = function() {
      if (noFixed || (storm.targetElement !== document.documentElement && storm.targetElement !== document.body)) {
        s.o.style.top = (screenY+scrollY-storm.flakeHeight)+'px';
      } else if (storm.flakeBottom) {
        s.o.style.top = storm.flakeBottom+'px';
      } else {
        s.o.style.display = 'none';
        s.o.style.top = 'auto';
        s.o.style.bottom = '0px';
        s.o.style.position = 'fixed';
        s.o.style.display = 'block';
      }
    };

    this.vCheck = function() {
      if (s.vX>=0 && s.vX<0.2) {
        s.vX = 0.2;
      } else if (s.vX<0 && s.vX>-0.2) {
        s.vX = -0.2;
      }
      if (s.vY>=0 && s.vY<0.2) {
        s.vY = 0.2;
      }
    };

    this.move = function() {
      var vX = s.vX*windOffset, yDiff;
      s.x += vX;
      s.y += (s.vY*s.vAmp);
      if (s.x >= screenX || screenX-s.x < storm.flakeWidth) { // X-axis scroll check
        s.x = 0;
      } else if (vX < 0 && s.x-storm.flakeLeftOffset < -storm.flakeWidth) {
        s.x = screenX-storm.flakeWidth-1; // flakeWidth;
      }
      s.refresh();
      yDiff = screenY+scrollY-s.y;
      if (yDiff<storm.flakeHeight) {
        s.active = 0;
        if (storm.snowStick) {
          s.stick();
        } else {
          s.recycle();
        }
      } else {
        if (storm.useMeltEffect && s.active && s.type < 3 && !s.melting && Math.random()>0.998) {
          // ~1/1000 chance of melting mid-air, with each frame
          s.melting = true;
          s.melt();
          // only incrementally melt one frame
          // s.melting = false;
        }
        if (storm.useTwinkleEffect) {
          if (!s.twinkleFrame) {
            if (Math.random()>0.9) {
              s.twinkleFrame = parseInt(Math.random()*20,10);
            }
          } else {
            s.twinkleFrame--;
            s.o.style.visibility = (s.twinkleFrame && s.twinkleFrame%2===0?'hidden':'visible');
          }
        }
      }
    };

    this.animate = function() {
      // main animation loop
      // move, check status, die etc.
      s.move();
    };

    this.setVelocities = function() {
      s.vX = vRndX+rnd(storm.vMaxX*0.12,0.1);
      s.vY = vRndY+rnd(storm.vMaxY*0.12,0.1);
    };

    this.setOpacity = function(o,opacity) {
      if (!opacitySupported) {
        return false;
      }
      o.style.opacity = opacity;
    };

    this.melt = function() {
      if (!storm.useMeltEffect || !s.melting) {
        s.recycle();
      } else {
        if (s.meltFrame < s.meltFrameCount) {
          s.meltFrame++;
          s.setOpacity(s.o,s.meltFrames[s.meltFrame]);
          s.o.style.fontSize = s.fontSize-(s.fontSize*(s.meltFrame/s.meltFrameCount))+'px';
          s.o.style.lineHeight = storm.flakeHeight+2+(storm.flakeHeight*0.75*(s.meltFrame/s.meltFrameCount))+'px';
        } else {
          s.recycle();
        }
      }
    };

    this.recycle = function() {
      s.o.style.display = 'none';
      s.o.style.position = (fixedForEverything?'fixed':'absolute');
      s.o.style.bottom = 'auto';
      s.setVelocities();
      s.vCheck();
      s.meltFrame = 0;
      s.melting = false;
      s.setOpacity(s.o,1);
      s.o.style.padding = '0px';
      s.o.style.margin = '0px';
      s.o.style.fontSize = s.fontSize+'px';
      s.o.style.lineHeight = (storm.flakeHeight+2)+'px';
      s.o.style.textAlign = 'center';
      s.o.style.verticalAlign = 'baseline';
      s.x = parseInt(rnd(screenX-storm.flakeWidth-20),10);
      s.y = parseInt(rnd(screenY)*-1,10)-storm.flakeHeight;
      s.refresh();
      s.o.style.display = 'block';
      s.active = 1;
    };

    this.recycle(); // set up x/y coords etc.
    this.refresh();

  };

  this.snow = function() {
    var active = 0, used = 0, waiting = 0, flake = null, i;
    for (i=s.flakes.length; i--;) {
      if (s.flakes[i].active === 1) {
        s.flakes[i].move();
        active++;
      } else if (s.flakes[i].active === 0) {
        used++;
      } else {
        waiting++;
      }
      if (s.flakes[i].melting) {
        s.flakes[i].melt();
      }
    }
    if (active<s.flakesMaxActive) {
      flake = s.flakes[parseInt(rnd(s.flakes.length),10)];
      if (flake.active === 0) {
        flake.melting = true;
      }
    }
  };

  this.mouseMove = function(e) {
    if (!s.followMouse) {
      return true;
    }
    var x = parseInt(e.clientX,10);
    if (x<screenX2) {
      windOffset = -windMultiplier+(x/screenX2*windMultiplier);
    } else {
      x -= screenX2;
      windOffset = (x/screenX2)*windMultiplier;
    }
  };

  this.createSnow = function(limit,allowInactive) {
    var i;
    for (i=0; i<limit; i++) {
      s.flakes[s.flakes.length] = new s.SnowFlake(s,parseInt(rnd(flakeTypes),10));
      if (allowInactive || i>s.flakesMaxActive) {
        s.flakes[s.flakes.length-1].active = -1;
      }
    }
    storm.targetElement.appendChild(docFrag);
  };

  this.timerInit = function() {
    s.timers = (!isWin98?[setInterval(s.snow,s.animationInterval)]:[setInterval(s.snow,s.animationInterval*3),setInterval(s.snow,s.animationInterval)]);
  };

  this.init = function() {
    var i;
    for (i=0; i<s.meltFrameCount; i++) {
      s.meltFrames.push(1-(i/s.meltFrameCount));
    }
    s.randomizeWind();
    s.createSnow(s.flakesMax); // create initial batch
    s.events.add(window,'resize',s.resizeHandler);
    s.events.add(window,'scroll',s.scrollHandler);
    if (s.freezeOnBlur) {
      if (isIE) {
        s.events.add(document,'focusout',s.freeze);
        s.events.add(document,'focusin',s.resume);
      } else {
        s.events.add(window,'blur',s.freeze);
        s.events.add(window,'focus',s.resume);
      }
    }
    s.resizeHandler();
    s.scrollHandler();
    if (s.followMouse) {
      s.events.add(isIE?document:window,'mousemove',s.mouseMove);
    }
    s.animationInterval = Math.max(20,s.animationInterval);
    s.timerInit();
  };

  this.start = function(bFromOnLoad) {
    if (!didInit) {
      didInit = true;
    } else if (bFromOnLoad) {
      // already loaded and running
      return true;
    }
    if (typeof s.targetElement === 'string') {
      var targetID = s.targetElement;
      s.targetElement = document.getElementById(targetID);
      if (!s.targetElement) {
        throw new Error('Snowstorm: Unable to get targetElement "'+targetID+'"');
      }
    }
    if (!s.targetElement) {
      s.targetElement = (!isIE?(document.documentElement?document.documentElement:document.body):document.body);
    }
    if (s.targetElement !== document.documentElement && s.targetElement !== document.body) {
      s.resizeHandler = s.resizeHandlerAlt; // re-map handler to get element instead of screen dimensions
    }
    s.resizeHandler(); // get bounding box elements
    s.usePositionFixed = (s.usePositionFixed && !noFixed); // whether or not position:fixed is supported
    fixedForEverything = s.usePositionFixed;
    if (screenX && screenY && !s.disabled) {
      s.init();
      s.active = true;
    }
  };

  function doDelayedStart() {
    window.setTimeout(function() {
      s.start(true);
    }, 20);
    // event cleanup
    s.events.remove(isIE?document:window,'mousemove',doDelayedStart);
  }

  function doStart() {
    if (!s.excludeMobile || !isMobile) {
      if (s.freezeOnBlur) {
        s.events.add(isIE?document:window,'mousemove',doDelayedStart);
      } else {
        doDelayedStart();
      }
    }
    // event cleanup
    s.events.remove(window, 'load', doStart);
  }

  // hooks for starting the snow
  s.events.add(window, 'load', doStart, false);

  return this;

}(window, document));

//-->