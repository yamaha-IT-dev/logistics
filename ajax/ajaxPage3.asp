<%@language="vbscript" codepage="1252"%>
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<%
Dim qry, iBrand
iBrand = ""
qry = Trim(Request.Querystring("qFunc"))
if qry = "jxBrand" then
  iBrand = Trim(Request.Querystring("qBrand"))
end if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head><title>Sample Autocomplete AJAX using Classic ASP</title>
<meta http-equiv="content-type" content="text/html; charset=us-ascii" />
<meta name="author" content="Rene Aquino Surop" />
<script type="text/javascript" src="ajaxPage3a.js"></script>
</head>
<body>
<form name="aspForm" action="ajaxPage3.asp" method="post" />
<div style="margin-top:15px; border:1px solid #3366ff; padding:10px;">
	<label for="jBrand">Car Brand: </label><div id="brandDiv" style="position:absolute; z-index:100;" /></div>&nbsp;
	<input id="txtBrand" name="txtBrand" value="<%=trim(iBrand)%>" onKeyUp="srchBrand(this.value)" />	
</div>
</form>
<div style="margin-top:15px; border:1px solid #3366ff; padding:10px;">
  The sample Autocomplete provides suggestions while you type into the field. Here the suggestions are brand names, displayed when a character are entered into the field.<br /><br />
  The datasource is a classic ASP server-side script which returns HTML data, specified via a simple URL for the source-option.
</div>
</body>
</html>