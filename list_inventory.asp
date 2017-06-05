<% response.cookies("current_URL_cookie_logistics") = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?" & Request.Querystring %>
<!--#include file="include/connection_it.asp " -->
<!--#include file="class/clsInventory.asp" -->
<% strSection = "inventory" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Inventory Master Maintenance</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="include/generic_form_validations.js"></script>
<script language="JavaScript" type="text/javascript">
function searchInventory(){
    var strPalletSearch = document.forms[0].txtSearch.value;
	var strDepartment 	= document.forms[0].cboDepartment.value;
	var strStatus 		= document.forms[0].cboStatus.value;

    document.location.href = 'list_inventory.asp?type=search&txtSearch=' + strPalletSearch + '&cboDepartment=' + strDepartment + '&cboStatus=' + strStatus;
}

function resetSearch(){
	document.location.href = 'list_inventory.asp?type=reset';
}

function validateUpdateInventoryForm(theForm) {
	var reason = "";
	var blnSubmit = true;

	reason += validateEmptyField(theForm.txtTariffCode,"Tariff Code");
	reason += validateSpecialCharacters(theForm.txtTariffCode,"Tariff Code");
	
	reason += validateEmptyField(theForm.txtTariffRate,"Tariff Rate");	
	reason += validateSpecialCharacters(theForm.txtTariffRate,"Tariff Rate");
	
	reason += validateEmptyField(theForm.cboDuty,"Duty");
	
  	if (reason != "") {
    	alert("Some fields need correction:\n" + reason);

		blnSubmit = false;
		return false;
  	}

	if (blnSubmit == true){
        theForm.action.value = 'update';
  		theForm.submit();

		return true;
    }
}
</script>
</head>
<body>
<%
sub setSearch
	select case Trim(Request("type"))
		case "reset"
			session("inventory_search") 		= ""
			session("inventory_department") 	= ""
			session("inventory_search_status") 	= ""
			session("inventory_initial_page") 	= 1
		case "search"
			session("inventory_search") 		= Trim(Request("txtSearch"))
			session("inventory_department") 	= Trim(Request("cboDepartment"))
			session("inventory_search_status") 	= Trim(Request("cboStatus"))
			session("inventory_initial_page") 	= 1
	end select
end sub

sub displayInventory
	dim iRecordCount
	iRecordCount = 0
    dim strPalletSearch
    dim strSQL
	dim strDamageSort
	dim strStatus
	dim intRecordCount
	dim strModifiedDate

	dim strTodayDate
	strTodayDate = FormatDateTime(Date())

	if session("inventory_search_status") = "" then
		session("inventory_search_status") = "2"
	end if

    call OpenWorkflowDataBase()

	set rs = Server.CreateObject("ADODB.recordset")

	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic
	rs.PageSize = 100

	strSQL = "SELECT * FROM workflow_inventory_master_reference "
	strSQL = strSQL & "	WHERE (base_code LIKE '%" & session("inventory_search") & "%' "
	strSQL = strSQL & "			OR item_type LIKE '%" & session("inventory_search") & "%' "
	strSQL = strSQL & "			OR creator_name LIKE '%" & session("inventory_search") & "%') "
	strSQL = strSQL & "		AND division LIKE '%" & session("inventory_department") & "%' "
	strSQL = strSQL & "		AND status_id LIKE '%" & session("inventory_search_status") & "%' "
	strSQL = strSQL & "	ORDER BY id"

	'Response.Write strSQL

	rs.Open strSQL, conn

	intPageCount = rs.PageCount
	intRecordCount = rs.recordcount
	
	Select Case Request("Action")
	    case "<<"
		    intpage = 1
			session("inventory_initial_page") = intpage
	    case "<"
		    intpage = Request("intpage") - 1
			session("inventory_initial_page") = intpage

			if session("inventory_initial_page") < 1 then session("inventory_initial_page") = 1
	    case ">"
		    intpage = Request("intpage") + 1
			session("inventory_initial_page") = intpage

			if session("inventory_initial_page") > intPageCount then session("inventory_initial_page") = IntPageCount
	    Case ">>"
		    intpage = intPageCount
			session("inventory_initial_page") = intpage
    end select

    strDisplayList = ""

	if not DB_RecSetIsEmpty(rs) Then

		For intRecord = 1 To rs.PageSize		
			'Form properties
			strDisplayList = strDisplayList & "<form method=""post"" name=""form_update_inventory"" id=""form_update_inventory"" onsubmit=""return validateUpdateInventoryForm(this)"">"
			strDisplayList = strDisplayList & "<input type=""hidden"" name=""action"" value=""update"">"
			strDisplayList = strDisplayList & "<input type=""hidden"" name=""id"" value=""" & trim(rs("id")) & """>"
			strDisplayList = strDisplayList & "<input type=""hidden"" name=""department"" value=""" & trim(rs("division")) & """>"
			strDisplayList = strDisplayList & "<input type=""hidden"" name=""type"" value=""" & trim(rs("item_type")) & """>"
			strDisplayList = strDisplayList & "<input type=""hidden"" name=""product"" value=""" & trim(rs("base_code")) & """>"			
			
			'Highlight updated
			if (DateDiff("d",rs("last_update_date"), strTodayDate) = 0) OR (DateDiff("d",rs("create_date"), strTodayDate) = 0) then
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
			
			strDisplayList = strDisplayList & "<td>" & FormatDateTime(rs("create_date"),2) & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("creator_name") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("id")
			if DateDiff("d",rs("create_date"), strTodayDate) = 0 then
				strDisplayList = strDisplayList & " <img src=""images/icon_new.gif"" border=""0"">"
			end if
			strDisplayList = strDisplayList & "</td>"
			
			strDisplayList = strDisplayList & "<td>" & rs("division") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("item_type") & "</td>"			
			strDisplayList = strDisplayList & "<td>" & rs("base_code") & "</td>"
			strDisplayList = strDisplayList & "<td><input type=""text"" id=""txtTariffCode"" name=""txtTariffCode"" maxlength=""20"" size=""20"" value=""" & rs("tariff_code") & """ ></td>"
			strDisplayList = strDisplayList & "<td><input type=""text"" id=""txtTariffRate"" name=""txtTariffRate"" maxlength=""15"" size=""15"" value=""" & rs("tariff_rate") & """ ></td>"
			strDisplayList = strDisplayList & "<td>"
			strDisplayList = strDisplayList & "	<select name=""cboDuty"">" 
			select case rs("duty")
				case "Duty"
					strDisplayList = strDisplayList & "<option value="""">...</option>"	
					strDisplayList = strDisplayList & "<option value=""Duty"" selected>Duty</option>"
					strDisplayList = strDisplayList & "<option value=""Duty Free"">Duty Free</option>"
					strDisplayList = strDisplayList & "<option value=""Concession"">Concession</option>"					
				case "Duty Free"
					strDisplayList = strDisplayList & "<option value="""">...</option>"	
					strDisplayList = strDisplayList & "<option value=""Duty"">Duty</option>"
					strDisplayList = strDisplayList & "<option value=""Duty Free"" selected>Duty Free</option>"
					strDisplayList = strDisplayList & "<option value=""Concession"">Concession</option>"					
				case "Concession"
					strDisplayList = strDisplayList & "<option value="""">...</option>"	
					strDisplayList = strDisplayList & "<option value=""Duty"">Duty</option>"
					strDisplayList = strDisplayList & "<option value=""Duty Free"">Duty Free</option>"
					strDisplayList = strDisplayList & "<option value=""Concession"" selected>Concession</option>"								
				case else
					strDisplayList = strDisplayList & "<option value="""">...</option>"
					strDisplayList = strDisplayList & "<option value=""Duty"">Duty</option>"
					strDisplayList = strDisplayList & "<option value=""Duty Free"">Duty Free</option>"
					strDisplayList = strDisplayList & "<option value=""Concession"">Concession</option>"					
			end select
			strDisplayList = strDisplayList & "	</select>"
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td>"
			strDisplayList = strDisplayList & "	<select name=""cboFTA"">" 
			select case rs("fta")
				case 0						
					strDisplayList = strDisplayList & "<option value=""0"" selected>No</option>"
					strDisplayList = strDisplayList & "<option value=""1"">Yes</option>"
				case 1					
					strDisplayList = strDisplayList & "<option value=""0"">No</option>"					
					strDisplayList = strDisplayList & "<option value=""1"" selected>Yes</option>"
			end select
			strDisplayList = strDisplayList & "	</select>"
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td><input type=""submit"" value=""Update"" /></td>"
			strDisplayList = strDisplayList & "<td>"
			Select Case	rs("status_id")
				case 2
					strDisplayList = strDisplayList & "Pending GM Approval"
				case 3
					strDisplayList = strDisplayList & "Pending Service Mgr Approval"
				case 4
					strDisplayList = strDisplayList & "Pending Logistics Approval"
				case 5
					strDisplayList = strDisplayList & "Completed"
				case 6
					strDisplayList = strDisplayList & "Cancelled"
			end select	
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td>" & FormatDateTime(rs("last_update_date"),2) & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("last_update_people") & "</td>"
			strDisplayList = strDisplayList & "</tr>"
			strDisplayList = strDisplayList & "</form>"
			rs.movenext
			iRecordCount = iRecordCount + 1
			If rs.EOF Then Exit For
		next

	else
        strDisplayList = "<tr class=""innerdoc""><td colspan=""14"" align=""center"">No records found.</td></tr>"
	end if

	strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
	strDisplayList = strDisplayList & "<td colspan=""14"" class=""recordspaging"">"
	strDisplayList = strDisplayList & "<form name=""MovePage"" action=""list_inventory.asp"" method=""post"">"
    strDisplayList = strDisplayList & "<input type=""hidden"" name=""intpage"" value=" & session("inventory_initial_page") & ">"

	if session("inventory_initial_page") = 1 then
   		strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value=""<<"">"
    	strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value=""<"">"
	else
		strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value=""<<"">"
    	strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value=""<"">"
	end if
	if session("inventory_initial_page") = intpagecount or intRecordCount = 0 then
    	strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value="">"">"
    	strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value="">>"">"
	else
		strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value="">"">"
    	strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value="">>"">"
	end if
    strDisplayList = strDisplayList & "<input type=""hidden"" name=""txtSearch"" value=" & strPalletSearch & ">"
	strDisplayList = strDisplayList & "<input type=""hidden"" name=""cboPalletType"" value=" & strPalletType & ">"
	strDisplayList = strDisplayList & "<br />"
    strDisplayList = strDisplayList & "Page: " & session("inventory_initial_page") & " to " & intpagecount
    strDisplayList = strDisplayList & "<br />"
	strDisplayList = strDisplayList & "<h2>Total: " & intRecordCount & " records.</h2>"
    strDisplayList = strDisplayList & "</form>"
    strDisplayList = strDisplayList & "</td>"
    strDisplayList = strDisplayList & "</tr>"

    call CloseDataBase()
end sub

sub main
	call UTL_validateLogin
	call setSearch

    if trim(session("inventory_initial_page"))  = "" then
    	session("inventory_initial_page") = 1
	end if    
	
	if Request.ServerVariables("REQUEST_METHOD") = "POST" then
		intID 			= Request("id")
		strDepartment	= Request("department")
		strType			= Request("type")
		strProduct		= Request("product")		
		strTariffCode	= Trim(Replace(Request.Form("txtTariffCode"),"'","''"))
		strTariffRate 	= Trim(Replace(Request.Form("txtTariffRate"),"'","''"))
		strDuty			= Trim(Request.Form("cboDuty"))
		intFTA			= Trim(Request.Form("cboFTA"))
		
		Select Case Trim(Request("action"))
			case "update"
				'response.write intID & "-" & strTariffCode & "-" & strTariffRate & "-" & strDuty & "-"
				call updateInventory(intID, strDepartment, strType, strProduct, strTariffCode, strTariffRate, strDuty, intFTA, session("UsrUserName"))
				'call displayInventory
		end select
	else
		call displayInventory
	end if	
end sub

call main

dim rs, intPageCount, intpage, intRecord, strDisplayList, strMessageText
dim intID, strDepartment, strType, strProduct, strTariffCode, strTariffRate, strDuty, intFTA
%>
<table cellspacing="0" cellpadding="0" align="center" class="full_size_table" border="0">
  <!-- #include file="include/header.asp" -->
  <tr>
    <td class="first_content">
    <h2>Inventory Master Maintenance for MOL (Prototype!)</h2>
    <div class="alert alert-search">
        <form name="frmSearch" id="frmSearch" action="list_inventory.asp?type=search" method="post" onsubmit="searchInventory()">
          <h3>Search Parameters:</h3>
          Product / Type / Created by:
          <input type="text" name="txtSearch" maxlength="25" size="25" value="<%= request("txtSearch") %>" />
          <select name="cboDepartment" onchange="searchInventory()">
            <option <% if session("inventory_department") = "" then Response.Write " selected" end if%> value="">All Departments</option>
            <option <% if session("inventory_department") = "AV" then Response.Write " selected" end if%> value="AV">AV</option>
            <option <% if session("inventory_department") = "MPD" then Response.Write " selected" end if%> value="MPD">MPD</option>
          </select>
          <select name="cboStatus" onchange="searchInventory()">
            <option <% if session("inventory_search_status") = "2" then Response.Write " selected" end if%> value="2">Open</option>
            <option <% if session("inventory_search_status") = "6" then Response.Write " selected" end if%> value="6">Cancelled</option>
          </select>
          <input type="button" name="btnSearch" value="Search" onclick="searchInventory()" />
          <input type="button" name="btnReset" value="Reset" onclick="resetSearch()" />
        </form>
      </div>
      <table cellspacing="0" cellpadding="8" class="database_records_nowidth" width="100%" border="0">
      <thead>
        <tr>
          <td width="10%">Created</td>
          <td width="8%">Requestor</td>
          <td width="5%">ID</td>
          <td width="5%">Dept</td>
          <td width="5%">Type</td>
          <td width="10%">Product</td>
          <td width="10%">Tariff Code</td>
          <td width="10%">Tariff Rate</td>
          <td width="10%">Duty</td>
          <td width="5%">FTA</td>
          <td width="2%">&nbsp;</td>
          <td width="10%">Status</td>
          <td width="10%">Last Updated Date</td>
          <td width="10%">By</td>
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