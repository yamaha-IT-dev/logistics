<% response.cookies("current_URL_cookie_logistics") = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?" & Request.Querystring %>
<!--#include file="include/connection_it.asp " -->
<!--#INCLUDE FILE = "include/AntiFixation.asp" -->
<% AntiFixationVerify("default.asp") %>
<% strSection = "shipment" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Shipments</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="include/javascript.js"></script>
<script language="JavaScript" type="text/javascript">
function searchShipment() {
    var strSearch       = document.forms[0].txtSearch.value;
    var strDepartment   = document.forms[0].cboDepartment.value;
    var strCountry      = document.forms[0].cboCountry.value;
    var strWarehouse    = document.forms[0].cboWarehouse.value;
    var strFTA          = document.forms[0].cboFTA.value;
    document.location.href = 'list_shipment.asp?type=search&txtSearch=' + strSearch + '&country=' + strCountry + '&department=' + strDepartment + '&warehouse=' + strWarehouse + '&fta=' + strFTA;
}

function resetSearch() {
    document.location.href = 'list_shipment.asp?type=reset';
}
</script>
</head>
<body>
<%
sub setSearch
    select case Trim(Request("type"))
        case "reset"
            session("shipment_search")      = ""
            session("shipment_department")  = ""
            session("shipment_country")     = ""
            session("shipment_warehouse")   = ""
            session("shipment_fta")         = ""
        case "search"
            session("shipment_search")      = Trim(Request("txtSearch"))
            session("shipment_department")  = Trim(Request("department"))
            session("shipment_country")     = Trim(Request("country"))
            session("shipment_warehouse")   = request("warehouse")
            session("shipment_fta")         = request("fta")
    end select
end sub

sub displayShipment
    dim iRecordCount
    iRecordCount = 0
    Dim strSortBy
    dim strSortItem
    Dim strSearch
    dim strSQL
    dim strDepartment
    dim strCountry

    dim strPageResultNumber
    dim strRecordPerPage
    dim intRecordCount
    dim strModifiedDate

    dim strTodayDate
    strTodayDate = FormatDateTime(Date())

    call OpenDataBase()

    set rs = Server.CreateObject("ADODB.recordset")

    rs.CursorLocation = 3   'adUseClient
    rs.CursorType = 3       'adOpenStatic

    strPageResultNumber = trim(Request("cboDealerResultSize"))
    strRecordPerPage = 900

    rs.PageSize = 900

    strSQL = "SELECT * FROM yma_shipment "
    strSQL = strSQL & " WHERE (supplier_invoice_no LIKE '%" & session("shipment_search") & "%' "
    strSQL = strSQL & "			OR container_no LIKE '%" & session("shipment_search") & "%' "
    strSQL = strSQL & "			OR EFT LIKE '%" & session("shipment_search") & "%' "
    strSQL = strSQL & "			OR vessel_name LIKE '%" & session("shipment_search") & "%') "
    strSQL = strSQL & "		AND department LIKE '%" & session("shipment_department") & "%' "
    strSQL = strSQL & "		AND country_origin LIKE '%" & session("shipment_country") & "%' "
    strSQL = strSQL & "		AND warehouse LIKE '%" & session("shipment_warehouse") & "%' "
    strSQL = strSQL & "		AND fta_status LIKE '%" & session("shipment_fta") & "%' "
    strSQL = strSQL & "		AND status = 1 "
    strSQL = strSQL & "	ORDER BY eta_unpacked DESC, eta_discharged DESC, supplier_invoice_no"

    rs.Open strSQL, conn

    'Response.Write strSQL

    intPageCount = rs.PageCount
    intRecordCount = rs.recordcount

    strDisplayList = ""

    if not DB_RecSetIsEmpty(rs) Then
        For intRecord = 1 To rs.PageSize

            if (DateDiff("d",rs("modified_date"), strTodayDate) = 0) OR (DateDiff("d",rs("date_created"), strTodayDate) = 0) then
                if iRecordCount Mod 2 = 0 then
                    strDisplayList = strDisplayList & "<tr class=""updated_today"">"
                else
                    strDisplayList = strDisplayList & "<tr class=""updated_today_2"">"
                end if
            else
                if iRecordCount Mod 2 = 0 then
                    strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
                else
                    strDisplayList = strDisplayList & "<tr class=""innerdoc_2"">"
                end if
            end if
            strDisplayList = strDisplayList & "<td><a href=""update_shipment.asp?ref=open&id=" & rs("shipment_id") & """>" & rs("shipment_id") & "</a></td>"

            strDisplayList = strDisplayList & "<td>" & rs("container_no") & "</td>"
            strDisplayList = strDisplayList & "<td>"

            Select Case Session("UsrLoginRole")
                case 3 'MOL
                    'strDisplayList = strDisplayList & "<a href=""ftp://203.221.101.249/shipment/" & rs("supplier_invoice_no") & """ target=""_blank"">" & rs("supplier_invoice_no") & "</a>"
					strDisplayList = strDisplayList & "<a href=""ftp://49.255.181.168/shipment/" & rs("supplier_invoice_no") & """ target=""_blank"">" & rs("supplier_invoice_no") & "</a>"
					
                case 14 'TT Warehouse Admin
                    strDisplayList = strDisplayList & "<a href=""file:\\172.29.67.6\shipment\" & rs("supplier_invoice_no") & """ target=""_blank"">" & rs("supplier_invoice_no") & "</a>"
                case 15 'TT Warehouse Normal Users
                    strDisplayList = strDisplayList & "<a href=""file:\\172.29.67.6\shipment\" & rs("supplier_invoice_no") & """ target=""_blank"">" & rs("supplier_invoice_no") & "</a>"
                case 16 'TT Office
                    strDisplayList = strDisplayList & "<a href=""ftp://yamaha_vic%5CTTLogShipment:ttL0gix@203.221.101.249/shipment/" & rs("supplier_invoice_no") & """ target=""_blank"">" & rs("supplier_invoice_no") & "</a>"
                case 17 
                    strDisplayList = strDisplayList & "<a href=""ftp://yamaha_vic%5CTTLogShipment:ttL0gix@203.221.101.249/shipment/" & rs("supplier_invoice_no") & """ target=""_blank"">" & rs("supplier_invoice_no") & "</a>"
                case else
                    strDisplayList = strDisplayList & "<a href=""file:\\YAMMAS22\shipment\" & rs("supplier_invoice_no") & """ target=""_blank"">" & rs("supplier_invoice_no") & "</a>"
            end select

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
            strDisplayList = strDisplayList & "<td>" & rs("department") & "</td>"
            strDisplayList = strDisplayList & "<td>"
            Select Case rs("transit")
                case 1
                    strDisplayList = strDisplayList & "<img src=""images/tick.gif"" border=""0"">"
                case 0
                    strDisplayList = strDisplayList & "<img src=""images/cross.gif"" border=""0"">"
            end select
            strDisplayList = strDisplayList & "</td>"

            strDisplayList = strDisplayList & "<td>"
            Select Case rs("custom_cleared")
                case "Y"
                    strDisplayList = strDisplayList & "<img src=""images/tick.gif"" border=""0"">"
                case "N"
                    strDisplayList = strDisplayList & "<img src=""images/cross.gif"" border=""0"">"
                case else
                    strDisplayList = strDisplayList & rs("custom_cleared")
            end select
            strDisplayList = strDisplayList & "</td>"
            strDisplayList = strDisplayList & "<td>"
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

            strDisplayList = strDisplayList & "<td>"
            Select Case rs("edo")
                case 0
                    strDisplayList = strDisplayList & "<img src=""images/cross.gif"" border=""0"">"
                case 1
                    strDisplayList = strDisplayList & "<img src=""images/tick.gif"" border=""0"">"
            end select
            strDisplayList = strDisplayList & "</td>"
            strDisplayList = strDisplayList & "<td>" & rs("EFT") & "</td>"

            strDisplayList = strDisplayList & "<td>"
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

            strDisplayList = strDisplayList & "<td>"
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

            strDisplayList = strDisplayList & "<td>" & rs("commodity") & "</td>"
            strDisplayList = strDisplayList & "<td>" & rs("port_origin") & "</td>"
            strDisplayList = strDisplayList & "<td>"
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
                case "USA"
                    strDisplayList = strDisplayList & "<img src=""images/usa.gif"" border=""0""> USA"
                case else
                    strDisplayList = strDisplayList & rs("country_origin")
            end select
            strDisplayList = strDisplayList & "</td>"
            strDisplayList = strDisplayList & "<td>" & rs("vessel_name") & "</td>"

            if rs("warehouse") = "TT" then
                strDisplayList = strDisplayList & "<td><img src=""images/tt.gif"" alt=""TT"" border=""0""></td>"
            else
                strDisplayList = strDisplayList & "<td>" & rs("warehouse") & "</td>"
            end if
            strDisplayList = strDisplayList & "<td>" & rs("cartons") & "</td>"
            if rs("date_shipment") = "01/01/1900" or rs("date_shipment") = "1/1/1900" then
                strDisplayList = strDisplayList & "<td class=""orange_text"">TBA</td>"
            else
                strDisplayList = strDisplayList & "<td nowrap>" & FormatDateTime(rs("date_shipment"),2) & "</td>"
            end if

            if rs("eta_discharged") = "01/01/1900" or rs("eta_discharged") = "1/1/1900" then
                strDisplayList = strDisplayList & "<td class=""orange_text"">TBA"
            else
                strDisplayList = strDisplayList & "<td nowrap>" & FormatDateTime(rs("eta_discharged"),2) & ""
            end if

            if rs("melb_time") <> "" then
                strDisplayList = strDisplayList & " - " & rs("melb_time") & " "
            end if
            strDisplayList = strDisplayList & "</td>"

            if rs("eta_availability") = "01/01/1900" or rs("eta_availability") = "1/1/1900" then
                strDisplayList = strDisplayList & "<td class=""orange_text"">TBA</td>"
            else
                strDisplayList = strDisplayList & "<td nowrap>" & FormatDateTime(rs("eta_availability"),2) & ""
            end if

            if rs("melb_eta_time") <> "" then
                strDisplayList = strDisplayList & " - " & rs("melb_eta_time") & " "
            end if
            strDisplayList = strDisplayList & "</td>"

            if rs("eta_unpacked") = "01/01/1900" or rs("eta_unpacked") = "1/1/1900" then
                strDisplayList = strDisplayList & "<td class=""orange_text"">TBA"
            else
                strDisplayList = strDisplayList & "<td nowrap>" & FormatDateTime(rs("eta_unpacked"),2) & ""
            end if

            if rs("unpack_time") <> "" then
                strDisplayList = strDisplayList & " - " & rs("unpack_time") & " "
            end if
            strDisplayList = strDisplayList & "</td>"

            strDisplayList = strDisplayList & "<td>" & rs("teu") & "</td>"
            strDisplayList = strDisplayList & "<td>"
            Select Case rs("paperwork")
                case 1
                    strDisplayList = strDisplayList & "<img src=""images/tick.gif"" border=""0"">"
                case 0
                    strDisplayList = strDisplayList & "<img src=""images/cross.gif"" border=""0"">"
            end select
            strDisplayList = strDisplayList & "</td>"
            strDisplayList = strDisplayList & "<td>" & rs("delivery_type") & "</td>"
            strDisplayList = strDisplayList & "<td><a onclick=""return confirm('Are you sure you want to delete " & rs("supplier_invoice_no") & " ?');"" href='delete_shipment.asp?id=" & rs("shipment_id") & "'><img src=""images/btn_delete.gif"" border=""0""></a></td>"
            strDisplayList = strDisplayList & "</tr>"

            rs.movenext
            iRecordCount = iRecordCount + 1
            If rs.EOF Then Exit For
        next
    else
        strDisplayList = "<tr class=""innerdoc""><td colspan=""25"" align=""center"">No shipments found.</td></tr>"
    end if

    strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
    strDisplayList = strDisplayList & "<td colspan=""25"" class=""recordspaging"">"
    strDisplayList = strDisplayList & "<input type=""hidden"" name=""txtSearch"" value=" & strSearch & ">"
    strDisplayList = strDisplayList & "<input type=""hidden"" name=""cboDepartment"" value=" & strDepartment & ">"
    strDisplayList = strDisplayList & "<input type=""hidden"" name=""cboCountry"" value=" & strCountry & ">"
    strDisplayList = strDisplayList & "<input type=""hidden"" name=""cboWarehouse"" value=" & strWarehouse & ">"
    strDisplayList = strDisplayList & "<h2>Search results: " & intRecordCount & " shipments.</h2>"
    strDisplayList = strDisplayList & "</td>"
    strDisplayList = strDisplayList & "</tr>"

    call CloseDataBase()
end sub

sub main
    call UTL_validateLogin
    call setSearch
    call displayShipment
end sub

call main

dim rs
dim intPageCount, intpage, intRecord
dim strDisplayList
%>
<table cellspacing="0" cellpadding="0" align="center" class="full_size_table" border="0">
  <!-- #include file="include/header.asp" -->
  <!-- #include file="list_shipment_forward.asp" -->
  <tr>
    <td class="first_content"><table cellpadding="5" cellspacing="0" border="0">
        <tr>
          <td valign="top"><img src="images/icon_shipment.jpg" border="0" alt="Shipment" /></td>
          <td valign="top"><div class="alert alert-success"><img src="images/add_icon.png" border="0" align="bottom" /> <a href="add_shipment.asp">Add Shipment</a></div>
            <p><img src="images/icon_excel.jpg" border="0" /> <a href="export_shipment.asp">Export</a></p></td>
          <td valign="top"><div class="alert alert-search">
              <form name="frmSearch" id="frmSearch" action="list_shipment.asp?type=search" method="post" onsubmit="searchShipment()">
                <h3>Search Parameters:</h3>
                Invoice no / Container no / EFT / Vessel name:
                <input type="text" name="txtSearch" size="25" value="<%= request("txtSearch") %>" maxlength="20" />
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
                <select name="cboWarehouse" onchange="searchShipment()">
                  <option value="">All Warehouses</option>
                  <option <% if session("shipment_warehouse") = "T" then Response.Write " selected" end if%> value="T">TT & 3T (m)</option>
                  <option <% if session("shipment_warehouse") = "3K" then Response.Write " selected" end if%> value="3K">3K & 3K (m)</option>
                </select>
                <select name="cboFTA" onchange="searchShipment()">
                  <option value="">All FTA</option>
                  <option <% if session("shipment_fta") = "1" then Response.Write " selected" end if%> value="1">Certificate of Origin</option>
                  <option <% if session("shipment_fta") = "2" then Response.Write " selected" end if%> value="2">Refund Application</option>
                  <option <% if session("shipment_fta") = "3" then Response.Write " selected" end if%> value="3">Import Declaration</option>
                  <option <% if session("shipment_fta") = "4" then Response.Write " selected" end if%> value="4">Refund (Complete)</option>
                </select>
                <input type="button" name="btnSearch" value="Search" onclick="searchShipment()" />
                <input type="button" name="btnReset" value="Reset" onclick="resetSearch()" />
              </form>
            </div></td>
        </tr>
      </table>
      <p><span class="current_header">Open</span> - <a href="list_past-shipment.asp">Past Shipments</a></p></td>
  </tr>
  <tr>
    <td><table cellspacing="0" cellpadding="8" class="database_records">
    <thead>
        <tr>
          <td>ID</td>
          <td>Container</td>
          <td>Invoice</td>
          <td>Dept</td>
          <td>Transit</td>
          <td>Custom</td>
          <td>Fumigation</td>
          <td>EDO</td>
          <td>EFT</td>
          <td>Doc</td>
          <td>FTA</td>
          <td>Commodity</td>
          <td>Port</td>
          <td>Country</td>
          <td>Vessel</td>
          <td>WH</td>
          <td>Cart</td>
          <td><span title="Date of Shipment">Shipment</span></td>
          <td><span title="Estimated Arrival Melbourne / Actual Delivery Time">Melb</span></td>
          <td><span title="Estimated Container Availability">Container</span></td>
          <td><span title="Estimated Unpack at Kagan or Base">Unpack</span></td>
          <td>TEU</td>
          <td><span title="Paperwork sent to Rocke">Rocke</span></td>
          <td>Delivery</td>
          <td>&nbsp;</td>
        </tr>
        </thead>
        <tbody>
        <%= strDisplayList %>
        </tbody>
      </table></td>
  </tr>
</table>
</body>
</html>