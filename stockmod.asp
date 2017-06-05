<% strSection = "services" %>
<!--#include file="include/connection_it.asp " -->
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>INTRANET - Yamaha Music Australia - Services - Stock Modifications and Approvals</title>
<link rel="stylesheet" href="../include/stylesheet.css">
<script src="../include/javascript.js"></script>
<script language="JavaScript" type="text/javascript">
function searchShipment(){    
    var strSearch = document.forms[0].txtSearch.value;
	var strType  = document.forms[0].cboType.value;
	var strSort  = document.forms[0].cboSort.value;
     document.location.href = 'stockmod.asp?type=search&txtSearch=' + strSearch + '&cboType=' + strType + '&cboSort=' + strSort;         
}
    
function resetSearch(){
	document.location.href = 'stockmod.asp?type=reset';    
}  
</script>
</head>
<body class="services_stock-modification" onLoad="xcSet('x','xc','co','main_services');">
<style type="text/css">
.services_stock-modification #services_stock-modification { font-weight:bold }
</style>
<%
sub setSearch	
	select case Trim(Request("type"))
		case "reset"
			session("strSearch") = ""
			session("strType") = ""
			session("strSort") = ""
		case "search"
			session("strSearch") = Trim(Request("txtSearch"))
			session("strType") = Trim(Request("cboType"))
			session("strSort") = Trim(Request("cboSort"))
	end select
end sub

sub displayShipment	
	dim iRecordCount
	iRecordCount = 0
    Dim strSortBy
	dim strSortItem
    Dim strSearch
    dim strSQL
	dim strType
	dim strSort
		
	dim strPageResultNumber
	dim strRecordPerPage
	dim intRecordCount
	dim strModifiedDate
	
	dim strTodayDate
	strTodayDate = FormatDateTime(Date())

	strSearch = ""
    if len(trim(Session("strSearch"))) > 0 then
        strSearch = Session("strSearch")
    end if
		
	strType = ""
    if len(trim(Session("strType"))) > 0 then
        strType = Session("strType")
    end if
	
	strSort = ""
    if len(trim(Session("strSort"))) > 0 then
        strSort = Session("strSort")
    end if
	
    call OpenDataBase()
	
	set rs = Server.CreateObject("ADODB.recordset")
	
	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic
	
	strPageResultNumber = trim(Request("cboDealerResultSize"))	
	strRecordPerPage = 900
			
	rs.PageSize = 900
	
	if Trim(Request("cboSort")) = "" then
		strSQL = "SELECT * FROM yma_stock_mod WHERE model_type LIKE '%" & Trim(Request("cboType")) & "%' AND status = 1 AND (model_name LIKE '%" & Trim(Request("txtSearch")) & "%' OR part_no_base LIKE '%" & Trim(Request("txtSearch")) & "%' OR created_by LIKE '%" & Trim(Request("txtSearch")) & "%' OR vendor_model_no LIKE '%" & Trim(Request("txtSearch")) & "%')"
	else
		strSQL = "SELECT * FROM yma_stock_mod WHERE model_type LIKE '%" & Trim(Request("cboType")) & "%' AND status = 1 AND (model_name LIKE '%" & Trim(Request("txtSearch")) & "%' OR part_no_base LIKE '%" & Trim(Request("txtSearch")) & "%' OR created_by LIKE '%" & Trim(Request("txtSearch")) & "%' OR vendor_model_no LIKE '%" & Trim(Request("txtSearch")) & "%') ORDER BY " & Trim(Request("cboSort"))
	end if
	
	rs.Open strSQL, conn
	
	'Response.Write strSQL
	
	intPageCount = rs.PageCount
	intRecordCount = rs.recordcount	

    strDisplayList = ""
	
	if not DB_RecSetIsEmpty(rs) Then
	
	    rs.AbsolutePage = session("cinitialPage")  
	
		For intRecord = 1 To rs.PageSize 
			'strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
			if (DateDiff("d",rs("date_modified"), strTodayDate) = 0) OR (DateDiff("d",rs("date_created"), strTodayDate) = 0) then
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
				
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("model_name")
			if DateDiff("d",rs("date_created"), strTodayDate) = 0 then
				strDisplayList = strDisplayList & " <img src=""images/icon_new.gif"" border=""0"">"
			end if
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("model_type") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("part_no_base") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("vendor_model_no") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">"
			if rs("hardwired") = 1 then
				strDisplayList = strDisplayList & "Yes"
			else
				strDisplayList = strDisplayList & "No"
			end if
			
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">"
			Select Case rs("document")
				case "1"
					strDisplayList = strDisplayList & "<a href=""file:\\YAMMAS22\shipment\_stockmods\" & rs("model_name") & ".doc"" target=""_blank"">View</a> "					
				case "0"
					strDisplayList = strDisplayList & "No"
				case else
					strDisplayList = strDisplayList & "..."
			end select
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("date_created") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("created_by") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("date_modified") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("modified_by") & "</td>"			
			strDisplayList = strDisplayList & "</tr>"
			
			rs.movenext
			iRecordCount = iRecordCount + 1	
			If rs.EOF Then Exit For 
		next

	else
        strDisplayList = "<tr class=innerdoc><td colspan=10 align=center>There is no stock modification.</td></tr>"
	end if
	
	strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
	strDisplayList = strDisplayList & "<td colspan=""10"" class=""recordspaging"">"
	'strDisplayList = strDisplayList & "<form name=""MovePage"" action=""list_stockmod.asp"" method=""post"">"    
    strDisplayList = strDisplayList & "<input type=""hidden"" name=""txtSearch"" value=" & strSearch & ">"
	strDisplayList = strDisplayList & "<input type=""hidden"" name=""cboType"" value=" & strType & ">"
	strDisplayList = strDisplayList & "<input type=""hidden"" name=""cboSort"" value=" & strSort & ">"
	'strDisplayList = strDisplayList & "<input type=""hidden"" name=""txtSource"" value=" & strSource & ">"
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
	'Response.Write "<br>Previous: " & Request.ServerVariables("HTTP_REFERER") & Request.Querystring
    if trim(session("cinitialPage"))  = "" then
    	session("cinitialPage") = 1
	end if		
    
    call displayShipment
    call setSearch

end sub

call main

dim rs
dim intPageCount, intpage, intRecord
dim strDisplayList
dim strDealerResultList
dim strStateList
dim strSalesManagerList
%>
<table border="0" cellspacing="0" cellpadding="0" class="main_table">
  <!-- #include file="include/header_VB.asp" -->
  <tr>
    <td class="outercontent" align="left"><table border="0">
        <tr>
          <td class="main_left"><!-- #include file="include/menu.asp" --></td>
          <td class="main_center_full"><table border="0" cellspacing="0" cellpadding="0" class="gradient-style-full">
              <tr>
                <td><p><strong>Stock Modifications</strong></p></td>
              </tr>
            </table>
            <p><a href="../">Home</a> <img src="../images/forward_arrow.gif" /> <a href="../Divisions/Service/">Service</a> <img src="../images/forward_arrow.gif" /> Stock Modifications</p>
            <form name="frmSearch" id="frmSearch">
              <p>Search Model Name / Part # / Model # / Created By :
                <input type="text" name="txtSearch" size="15" value="<%= request("txtSearch") %>" />
                <select name="cboType">
                  <option value="">All Types</option>
                  <option <% if session("strType") = "LEAD" then Response.Write " selected" end if%> value="LEAD">LEAD</option>
                  <option <% if session("strType") = "PLUG" then Response.Write " selected" end if%> value="PLUG">PLUG</option>
                  <option <% if session("strType") = "ADAPTOR" then Response.Write " selected" end if%> value="ADAPTOR">ADAPTOR</option>
                  <option <% if session("strType") = "DVD" then Response.Write " selected" end if%> value="DVD">DVD</option>
                </select>
                <select name="cboSort">
                  <option value="">Sort by...</option>
                  <option <% if session("strSort") = "model_name" then Response.Write " selected" end if%> value="model_name">Model Name</option>
                  <option <% if session("strSort") = "model_type" then Response.Write " selected" end if%> value="model_type">Model Type</option>
                  <option <% if session("strSort") = "part_no_base" then Response.Write " selected" end if%> value="part_no_base">Part No</option>
                  <option <% if session("strSort") = "vendor_model_no" then Response.Write " selected" end if%> value="vendor_model_no">Vendor Model No</option>
                  <option <% if session("strSort") = "date_modified DESC" then Response.Write " selected" end if%> value="date_modified DESC">Last Modified Date</option>
                </select>
                <input type="button" name="btnSearch" value="Search" onclick="searchShipment()" />
                <input type="button" name="btnReset" value="Reset" onclick="resetSearch()" />
            </form>
            <table cellspacing="0" cellpadding="4" class="database_records">
              <tr class="innerdoctitle" align="center">
                <td width="15%"><b>Model Name</b></td>
                <td width="10%"><b>Model Type</b></td>
                <td width="15%"><b>Part # BASE</b></td>
                <td width="10%"><b>Vendor Model #</b></td>
                <td width="5%"><b>Hardwired</b></td>
                <td width="5%"><b>Document</b></td>
                <td width="10%"><b>Created Date</b></td>
                <td width="10%"><b>Created By</b></td>
                <td width="10%"><b>Last Modified Date</b></td>
                <td width="10%"><b>Last Modified By</b></td>
              </tr>
              <%= strDisplayList %>
            </table></td>
        </tr>        
      </table></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
</body>
</html>