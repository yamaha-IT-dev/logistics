<!--#include file="include/connection_it.asp " -->
<% strSection = "shipment" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Updated Shipments Today</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="include/javascript.js"></script>
<script language="JavaScript" type="text/javascript">
function searchShipment(){    
    var strSearch = document.forms[0].txtSearch.value;
	var strDepartment  = document.forms[0].cboDepartment.value;
	var strCountry  = document.forms[0].cboCountry.value;

	document.location.href = 'shipment_today.asp?type=search&txtSearch=' + strSearch + '&cboCountry=' + strCountry + '&cboDepartment=' + strDepartment;     	     
}
    
function resetSearch(){
	document.location.href = 'shipment_today.asp?type=reset';    
}  
</script>
</head>
<body>
<%
session.lcid = 2057

Function FormatMediumDate(DateValue)
    Dim strYYYY
    Dim strMM
    Dim strDD

    strYYYY = CStr(DatePart("yyyy", DateValue))

    strMM = CStr(DatePart("m", DateValue))
    If Len(strMM) = 1 Then strMM = "0" & strMM

    strDD = CStr(DatePart("d", DateValue))
    If Len(strDD) = 1 Then strDD = "0" & strDD

    FormatMediumDate = strMM & "/" & strDD & "/" & strYYYY
End Function

sub setSearch	
	select case Trim(Request("type"))
		case "reset"
			session("shipment_department") = ""
			session("shipment_country") = ""
		case "search"
			session("shipment_department") = request("cboDepartment")
			session("shipment_country") = request("cboCountry")
	end select
end sub

sub displayShipment	
	dim iRecordCount
	iRecordCount = 0
    Dim strSortBy
	dim strSortItem
    Dim strSearchTxt
    dim strSQL
	
	dim strPageResultNumber
	dim strRecordPerPage
	dim intRecordCount
	dim intStatus
	
	dim strTodayDate
	strTodayDate = Date()
	
	strSearchTxt = trim(Request("txtSearch"))

    call OpenDataBase()
	
	set rs = Server.CreateObject("ADODB.recordset")
	
	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic
	
	strPageResultNumber = trim(Request("cboDealerResultSize"))	
	strRecordPerPage = 900
			
	rs.PageSize = 900
	
	session("shipment_department") = trim(Request("cboDepartment"))
	session("shipment_country") = trim(Request("cboCountry"))	
	
	strSQL = "SELECT * FROM yma_shipment "
	strSQL = strSQL & "	WHERE modified_date > '" & FormatMediumDate(strTodayDate) & "' "
	strSQL = strSQL & "		AND department LIKE '%" & Request("cboDepartment") & "%' "
	strSQL = strSQL & "		AND country_origin LIKE '%" & Request("cboCountry") & "%' "
	strSQL = strSQL & "		AND (supplier_invoice_no LIKE '%" & trim(Request("txtSearch")) & "%' "
	strSQL = strSQL & "			OR container_no LIKE '%" & trim(Request("txtSearch")) & "%' "
	strSQL = strSQL & "			OR warehouse LIKE '%" & Trim(Request("txtSearch")) & "%' "
	strSQL = strSQL & "			OR EFT LIKE '%" & trim(Request("txtSearch")) & "%' "
	strSQL = strSQL & "			OR vessel_name LIKE '%" & trim(Request("txtSearch")) & "%') "
	strSQL = strSQL & "	ORDER BY eta_unpacked DESC, eta_discharged DESC, supplier_invoice_no"
	
	rs.Open strSQL, conn
	
	'Response.Write strSQL
	
	intPageCount = rs.PageCount
	intRecordCount = rs.recordcount	

    strDisplayList = ""
	
	if not DB_RecSetIsEmpty(rs) Then	
	
		For intRecord = 1 To rs.PageSize 

		'if (DateDiff("d",rs("modified_date"), strTodayDate) = 0) OR (DateDiff("d",rs("date_created"), strTodayDate) = 0) then
			'strDisplayList = strDisplayList & "<tr class=""updated_today"">"
			if iRecordCount Mod 2 = 0 then
				strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
			else
				strDisplayList = strDisplayList & "<tr class=""innerdoc_2"">"
			end if
					
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("container_no") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center""><a href=""file:\\YAMMAS22\shipment\" & rs("supplier_invoice_no") & """ target=""_blank"">" & rs("supplier_invoice_no") & "</a>"
			if DateDiff("d",rs("date_created"), strTodayDate) = 0 then
				strDisplayList = strDisplayList & " <img src=""images/icon_new.gif"" border=""0"">"
			end if
			if rs("air_freight") = 1 then
				strDisplayList = strDisplayList & " <img src=""images/airplane.gif"" border=""0"">"
			end if
			if rs("priority") = 1 then
				strDisplayList = strDisplayList & " <img src=""images/icon_priority.gif"" border=""0"">"
			end if
			strDisplayList = strDisplayList & "</td>"
			
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("department") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">"
			Select Case rs("custom_cleared")
				case "Y"
					strDisplayList = strDisplayList & "<img src=""images/tick.gif"" border=""0"">"
				case "N"
					strDisplayList = strDisplayList & "<img src=""images/cross.gif"" border=""0"">"					
				case else
			 		strDisplayList = strDisplayList & rs("custom_cleared")
			end select
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">"
			Select Case rs("fumigation")
				case "Y"
					strDisplayList = strDisplayList & "<img src=""images/tick.gif"" border=""0"">"
				case "N"
					strDisplayList = strDisplayList & "<img src=""images/cross.gif"" border=""0"">"
				case "-"
					strDisplayList = strDisplayList & "-"
				case else
			 		strDisplayList = strDisplayList & rs("fumigation")
			end select
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">"
			Select Case rs("edo")
				case 0
					strDisplayList = strDisplayList & "<img src=""images/cross.gif"" border=""0"">"
				case 1
					strDisplayList = strDisplayList & "<img src=""images/tick.gif"" border=""0"">"				
			end select
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("EFT") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">"
			Select Case rs("all_documents")
				case "Y"
					strDisplayList = strDisplayList & "<img src=""images/tick.gif"" border=""0"">"
				case "N"
					strDisplayList = strDisplayList & "<img src=""images/cross.gif"" border=""0"">"
				case "Part"
					strDisplayList = strDisplayList & "<font color=red><strong>PART</strong></font>"
				case else
			 		strDisplayList = strDisplayList & rs("all_documents")
			end select
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">"
			Select Case rs("fta_status")
				case "0"
					strDisplayList = strDisplayList & "-"
				case "1"
					strDisplayList = strDisplayList & "<img src=""images/bullet_certificate-origin.gif"" border=""0"">"
				case "2"
					strDisplayList = strDisplayList & "<img src=""images/bullet_refund-application.gif"" border=""0"">"
				case "3"
					strDisplayList = strDisplayList & "<img src=""images/bullet_import-declaration.gif"" border=""0"">"
				case "4"
					strDisplayList = strDisplayList & "<img src=""images/tick.gif"" border=""0"">"				
			end select
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("commodity") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("port_origin") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">"
			Select Case rs("country_origin")
				case "Indonesia"
					strDisplayList = strDisplayList & "<img src=""images/indo.gif"" border=""0""> INA"
				case "China"
					strDisplayList = strDisplayList & "<img src=""images/china.gif"" border=""0""> CHN"
				case "Malaysia"
					strDisplayList = strDisplayList & "<img src=""images/malaysia.gif"" border=""0""> MAL"
				case "Japan"
					strDisplayList = strDisplayList & "<img src=""images/japan.gif"" border=""0""> JPN"
				case "Vietnam"
					strDisplayList = strDisplayList & "<img src=""images/vietnam.gif"" border=""0""> VIE"
				case "NZ"
					strDisplayList = strDisplayList & "<img src=""images/nz.gif"" border=""0""> NZL"
				case else
			 		strDisplayList = strDisplayList & rs("country_origin")
			end select
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("vessel_name") & "</td>"
			'strDisplayList = strDisplayList & "<td align=""center"">" & rs("voyage") & "</td>"
			if rs("warehouse") = "TT" then
				strDisplayList = strDisplayList & "<td align=""center""><img src=""images/tt.gif"" alt=""TT"" border=""0""></td>"
			else
			 	strDisplayList = strDisplayList & "<td align=""center"">" & rs("warehouse") & "</td>"
			end if
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("cartons") & "</td>"
			if rs("date_shipment") = "01/01/1900" or rs("date_shipment") = "1/1/1900" then 
				strDisplayList = strDisplayList & "<td class=""orange_text"">TBA</td>"	
			else
				strDisplayList = strDisplayList & "<td align=""center"">" & FormatDateTime(rs("date_shipment"),1) & "</td>"	
			end if
			
			if rs("eta_discharged") = "01/01/1900" or rs("eta_discharged") = "1/1/1900" then 
				strDisplayList = strDisplayList & "<td class=""orange_text"">TBA</td>"	
			else
				strDisplayList = strDisplayList & "<td align=""center"">" & FormatDateTime(rs("eta_discharged"),1) & "</td>"
			end if						
			
			if rs("eta_availability") = "01/01/1900" or rs("eta_availability") = "1/1/1900" then 
				strDisplayList = strDisplayList & "<td class=""orange_text"">TBA</td>"	
			else
				strDisplayList = strDisplayList & "<td align=""center"">" & FormatDateTime(rs("eta_availability"),1) & ""	
			end if	
			
			if rs("melb_eta_time") <> "" then
				strDisplayList = strDisplayList & " - " & rs("melb_eta_time") & " "
			end if
			strDisplayList = strDisplayList & "</td>"	
			
			if rs("eta_unpacked") = "01/01/1900" or rs("eta_unpacked") = "1/1/1900" then 
				strDisplayList = strDisplayList & "<td class=""orange_text"">TBA</td>"	
			else
				strDisplayList = strDisplayList & "<td align=""center"">" & FormatDateTime(rs("eta_unpacked"),1) & "</td>"	
			end if
					
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("teu") & "</td>"
			if rs("status") = 1 then 
				strDisplayList = strDisplayList & "<td align=""center"">Open</td>"
			else
				strDisplayList = strDisplayList & "<td class=""green_text"">Completed</td>"
			end if 
			
			strDisplayList = strDisplayList & "</tr>"
			'end if
			rs.movenext
			iRecordCount = iRecordCount + 1
			If rs.EOF Then Exit For 
		next

	else
        strDisplayList = "<tr class=""innerdoc""><td colspan=""21"" align=""center"">There are no updated shipment records today.</td></tr>"
	end if
	
	strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
	strDisplayList = strDisplayList & "<td colspan=""21"" class=""recordspaging"">"
	'strDisplayList = strDisplayList & "<form name=""MovePage"" action=""shipment_today.asp"" method=""post"">"    
    strDisplayList = strDisplayList & "<input type=""hidden"" name=""txtSearch"" value=" & strSearchTxt & ">"
    strDisplayList = strDisplayList & "<br />"
	strDisplayList = strDisplayList & "Search results: " & intRecordCount & " records."
    'strDisplayList = strDisplayList & "</form>"
    strDisplayList = strDisplayList & "</td>"
    strDisplayList = strDisplayList & "</tr>"
	
	rs.close
	set rs = nothing
    call CloseDataBase()
end sub

sub main 
	if trim(session("strStatus"))  = "" then
    	session("strStatus") = 1
	end if
    
    call displayShipment
    call setSearch

end sub

call main

dim rs
dim intPageCount, intpage, intRecord
dim strDisplayList
%>
<table cellspacing="0" cellpadding="0" align="center" width="100%">
  <tr>
    <td align="center"><table cellpadding="8" cellspacing="0" border="0">
        <tr>
          <td valign="top"><img src="images/icon_shipment.jpg" border="0" /></td>
          <td><div class="alert alert-search">
              <form name="frmSearch" id="frmSearch" action="shipment_today.asp?type=search" method="post" onsubmit="searchShipment()">
                <h3>Search Parameters:</h3>
                Invoice no / Container no / EFT / Vessel / Warehouse :
                <input type="text" name="txtSearch" size="15" value="<%= request("txtSearch") %>" />
                <select name="cboDepartment" onchange="searchShipment()">
                  <option value="">All Departments</option>
                  <option <% if session("shipment_department") = "AV" then Response.Write " selected" end if%> value="AV">AV</option>
                  <option <% if session("shipment_department") = "MPD" then Response.Write " selected" end if%> value="MPD">MPD</option>
                  <option <% if session("shipment_department") = "AV-MPD" then Response.Write " selected" end if%> value="AV-MPD">AV & MPD</option>
                  <option <% if session("shipment_department") = "Other" then Response.Write " selected" end if%> value="Other">Other</option>
                </select>
                <select name="cboCountry" onchange="searchShipment()">
                  <option value="">All Countries</option>
                  <option <% if session("shipment_country") = "China" then Response.Write " selected" end if%> value="China">China</option>
                  <option <% if session("shipment_country") = "England" then Response.Write " selected" end if%> value="England">England</option>
                  <option <% if session("shipment_country") = "Germany" then Response.Write " selected" end if%> value="Germany">Germany</option>
                  <option <% if session("shipment_country") = "Indonesia" then Response.Write " selected" end if%> value="Indonesia">Indonesia</option>
                  <option <% if session("shipment_country") = "Japan" then Response.Write " selected" end if%> value="Japan">Japan</option>
                  <option <% if session("shipment_country") = "Malaysia" then Response.Write " selected" end if%> value="Malaysia">Malaysia</option>
                  <option <% if session("shipment_country") = "NZ" then Response.Write " selected" end if%> value="NZ">NZ</option>
                  <option <% if session("shipment_country") = "Singapore" then Response.Write " selected" end if%> value="Singapore">Singapore</option>
                  <option <% if session("shipment_country") = "USA" then Response.Write " selected" end if%> value="USA">USA</option>
                  <option <% if session("shipment_country") = "Vietnam" then Response.Write " selected" end if%> value="Vietnam">Vietnam</option>
                  <option <% if session("shipment_country") = "Other" then Response.Write " selected" end if%> value="Other">Other</option>
                </select>
                <input type="button" name="btnSearch" value="Search" onclick="searchShipment()" />
                <input type="button" name="btnReset" value="Reset" onclick="resetSearch()" />
              </form>
            </div></td>
          <td valign="top"><img src="images/yamaha_logo.jpg" border="0" /></td>
        </tr>
      </table>
      <h3><a href="shipment.asp">Open</a> | Updated Today | <a href="shipments_all.asp">All Shipment</a></h3>
      <div align="right"><a href="javascript:PrintThisPage()"><img src="../images/icon_printer.gif" alt="Printer Friendly" border="0" /></a></div>
      <br />
      <table cellspacing="0" cellpadding="4" class="database_records">
        <tr class="innerdoctitle" align="center">
          <td>Container</td>
          <td>Supplier invoice</td>
          <td>Dept</td>
          <td>Custom cleared</td>
          <td>Fumigation</td>
          <td>EDO</td>
          <td>EFT</td>
          <td>All docs</td>
          <td>FTA</td>
          <td>Commodity</td>
          <td>Port origin</td>
          <td>Country</td>
          <td>Vessel</td>
          <td>Warehouse</td>
          <td>No of carton</td>
          <td><span title="Date of Shipment">0 - Shipment Date</span></td>
          <td><span title="Estimated Arrival Melbourne">1 - Melb ETA</span></td>
          <td><span title="Estimated Container Availability">2 - Container ETA</span></td>
          <td><span title="Estimated Unpack at Kagan or Base">3 - Unpack ETA</span></td>
          <td>TEU</td>
          <td>Status</td>
        </tr>
        <%= strDisplayList %>
      </table></td>
  </tr>
</table>
</body>
</html>