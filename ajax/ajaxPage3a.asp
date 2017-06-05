<%
Dim action, iBrand, iResult
action = Trim(Request.QueryString("qAction"))
if action = "Search" then
  iBrand = Trim(Request.QueryString("qBrand"))
	'this part is actually a server-side Cobol COM wrapper or object that is fetching data as below;
	'DUMMY RESULT ONLY
	iResult = "<div style='border:1px #ccc solid; margin-top:23px; margin-left:5px; width:200px;'><table width='100%' cellpadding='0px' cellspacing='0px' align='center'>"
	iResult = iResult + "<tr bgcolor='#cccc99'><td style='width:100px'><a href='ajaxPage3.asp?qFunc=jxBrand&qBrand=Chevrolet'>Chevrolet</a></td></tr>"
	iResult = iResult + "<tr bgcolor='#ffffff'><td style='width:100px'><a href='ajaxPage3.asp?qFunc=jxBrand&qBrand=Ford'>Ford</a></td></tr>"
	iResult = iResult + "<tr bgcolor='#cccc99'><td style='width:100px'><a href='ajaxPage3.asp?qFunc=jxBrand&qBrand=Honda'>Honda</a></td></tr>"
	iResult = iResult + "<tr bgcolor='#ffffff'><td style='width:100px'><a href='ajaxPage3.asp?qFunc=jxBrand&qBrand=Hyundai'>Hyundai</a></td></tr>"
	iResult = iResult + "<tr bgcolor='#cccc99'><td style='width:100px'><a href='ajaxPage3.asp?qFunc=jxBrand&qBrand=Toyota'>Toyota</a></td></tr>"
	iResult = iResult + "</table></div>"

	Response.Write(iResult)
end if
%>