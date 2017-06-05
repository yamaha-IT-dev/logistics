<% response.cookies("current_URL_cookie_logistics") = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?" & Request.Querystring %>
<!--#include file="include/connection_it.asp " -->
<!--#include file="class/clsPallet.asp" -->
<% strSection = "gra" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Pallets</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="include/generic_form_validations.js"></script>
<script language="JavaScript" type="text/javascript">
function searchPallets(){
    var strPalletSearch = document.forms[0].txtSearch.value;
	var strDepartment 	= document.forms[0].cboDepartment.value;
	var strStatus 		= document.forms[0].cboStatus.value;

    document.location.href = 'list_pallet.asp?type=search&txtSearch=' + strPalletSearch + '&cboDepartment=' + strDepartment + '&cboStatus=' + strStatus;
}

function resetSearch(){
	document.location.href = 'list_pallet.asp?type=reset';
}

function validateAddPalletForm(theForm) {
	var reason 		= "";
	var blnSubmit 	= true;
	
	reason += validateEmptyField(theForm.cboPalletDepartment,"Department");	
	reason += validateSpecialCharacters(theForm.txtPalletInfo,"Pallet info");
	
  	if (reason != "") {
    	alert("Some fields need correction:\n" + reason);
    	
		blnSubmit = false;
		return false;
  	}

	if (blnSubmit == true){
        theForm.Action.value = 'add';
		
		return true;
    }
}

function validateUpdatePalletForm(theForm) {
	var reason = "";
	var blnSubmit = true;

	reason += validateEmptyField(theForm.txtUpdateDepartment,"Department");	
	reason += validateSpecialCharacters(theForm.txtUpdateInfo,"Pallet info");
	
  	if (reason != "") {
    	alert("Some fields need correction:\n" + reason);

		blnSubmit = false;
		return false;
  	}

	if (blnSubmit == true){
		//alert("updatezzz");
        theForm.Action.value = 'update';
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
			session("pallet_search") 		= ""
			session("pallet_department") 	= ""
			session("pallet_search_status") = ""
			session("pallet_initial_page") 	= 1
		case "search"
			session("pallet_search") 		= Trim(Request("txtSearch"))
			session("pallet_department") 	= Trim(Request("cboDepartment"))
			session("pallet_search_status") = Trim(Request("cboStatus"))
			session("pallet_initial_page") 	= 1
	end select
end sub



sub displayPallets
	dim iRecordCount
	iRecordCount = 0
    dim strDamageSortBy
	dim strDamageSortItem
    dim strPalletSearch
    dim strSQL
	'dim strPalletType
	dim strDamageSort
	dim strStatus

	dim strPageResultNumber
	dim strRecordPerPage
	dim intRecordCount
	dim strModifiedDate

	dim strTodayDate
	strTodayDate = FormatDateTime(Date())

	if session("pallet_search_status") = "" then
		session("pallet_search_status") = "1"
	end if

    call OpenDataBase()

	set rs = Server.CreateObject("ADODB.recordset")

	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic
	rs.PageSize = 1000

	strSQL = "SELECT * FROM tbl_pallets "
	strSQL = strSQL & "	WHERE (pallet_id LIKE '%" & session("pallet_search") & "%' "
	strSQL = strSQL & "			OR pallet_created_by LIKE '%" & session("pallet_search") & "%') "
	strSQL = strSQL & "		AND pallet_department LIKE '%" & session("pallet_department") & "%' "
	strSQL = strSQL & "		AND pallet_status LIKE '%" & session("pallet_search_status") & "%' "
	strSQL = strSQL & "	ORDER BY pallet_id DESC"

	'Response.Write strSQL

	rs.Open strSQL, conn

	intPageCount = rs.PageCount
	intRecordCount = rs.recordcount
	
	Select Case Request("Action")
	    case "<<"
		    intpage = 1
			session("pallet_initial_page") = intpage
	    case "<"
		    intpage = Request("intpage") - 1
			session("pallet_initial_page") = intpage

			if session("pallet_initial_page") < 1 then session("pallet_initial_page") = 1
	    case ">"
		    intpage = Request("intpage") + 1
			session("pallet_initial_page") = intpage

			if session("pallet_initial_page") > intPageCount then session("pallet_initial_page") = IntPageCount
	    Case ">>"
		    intpage = intPageCount
			session("pallet_initial_page") = intpage
    end select

    strDisplayList = ""

	if not DB_RecSetIsEmpty(rs) Then

		For intRecord = 1 To rs.PageSize
		
			'Form properties
			strDisplayList = strDisplayList & "<form method=""post"" name=""form_update_pallet"" id=""form_update_pallet"" onsubmit=""return validateUpdatePalletForm(this)"">"
			strDisplayList = strDisplayList & "<input type=""hidden"" name=""action"" value=""update"">"
			strDisplayList = strDisplayList & "<input type=""hidden"" name=""pallet_id"" value=""" & trim(rs("pallet_id")) & """>"
			
			'Highlight updated
			if (DateDiff("d",rs("pallet_date_modified"), strTodayDate) = 0) OR (DateDiff("d",rs("pallet_date_created"), strTodayDate) = 0) then
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
			
			
			
			
			
			'Pallet No
			strDisplayList = strDisplayList & "<td><a href=""view_pallet.asp?pallet_no=" & rs("pallet_id") & """>"
			strDisplayList = strDisplayList & "" & rs("pallet_id") & "</a>"
			if DateDiff("d",rs("pallet_date_created"), strTodayDate) = 0 then
				strDisplayList = strDisplayList & " <img src=""images/icon_new.gif"" border=""0"">"
			end if
			strDisplayList = strDisplayList & "</td>"
			
			'Department
			strDisplayList = strDisplayList & "<td><input type=""text"" id=""txtUpdateDepartment"" name=""txtUpdateDepartment"" maxlength=""5"" size=""8"" value=""" & rs("pallet_department") & """ ></td>"
			
			'Pallet Info
			strDisplayList = strDisplayList & "<td><input type=""text"" id=""txtUpdateInfo"" name=""txtUpdateInfo"" maxlength=""60"" size=""40"" value=""" & rs("pallet_info") & """ ></td>"
			
			'Pallet Status
			strDisplayList = strDisplayList & "<td><select name=""cboUpdateStatus"">"            
			if rs("pallet_status") = 1 then
				strDisplayList = strDisplayList & "<option value=""1"" selected>Open</option><option value=""0"">Completed</option>"
			else
				strDisplayList = strDisplayList & "<option value=""1"">Open</option><option value=""0"" selected>Completed</option>"
			end if
			strDisplayList = strDisplayList & " </select>"
			strDisplayList = strDisplayList & "</td>"			
			
			'Update button
			strDisplayList = strDisplayList & "<td><input type=""submit"" value=""Update"" /></td>"
			
			'Created by
			strDisplayList = strDisplayList & "<td>" & rs("pallet_created_by") & " "
			if IsNull(rs("pallet_date_created")) then
				strDisplayList = strDisplayList & "NA</td>"
			else
				strDisplayList = strDisplayList & " - " & FormatDateTime(rs("pallet_date_created"),2) & "</td>"
			end if
			
			'Modified by
			strDisplayList = strDisplayList & "<td>" & rs("pallet_modified_by") & " "
			if IsNull(rs("pallet_date_modified")) then
				strDisplayList = strDisplayList & "NA</td>"
			else
				strDisplayList = strDisplayList & " - " & FormatDateTime(rs("pallet_date_modified"),2) & "</td>"
			end if
			
			'Delete button
			strDisplayList = strDisplayList & "<td><a onclick=""return confirm('Are you sure you want to delete " & rs("pallet_id") & " ?');"" href='delete_pallet.asp?pallet_id=" & rs("pallet_id") & "'><img src=""images/btn_delete.gif"" border=""0""></a></td>"
			strDisplayList = strDisplayList & "</tr>"
			strDisplayList = strDisplayList & "</form>"
			rs.movenext
			iRecordCount = iRecordCount + 1
			If rs.EOF Then Exit For
		next

	else
        strDisplayList = "<tr class=""innerdoc""><td colspan=""9"" align=""center"">No pallets found.</td></tr>"
	end if

	strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
	strDisplayList = strDisplayList & "<td colspan=""9"" class=""recordspaging"">"
	strDisplayList = strDisplayList & "<form name=""MovePage"" action=""list_pallet.asp"" method=""post"">"
    strDisplayList = strDisplayList & "<input type=""hidden"" name=""intpage"" value=" & session("pallet_initial_page") & ">"

	if session("pallet_initial_page") = 1 then
   		strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value=""<<"">"
    	strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value=""<"">"
	else
		strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value=""<<"">"
    	strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value=""<"">"
	end if
	if session("pallet_initial_page") = intpagecount or intRecordCount = 0 then
    	strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value="">"">"
    	strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value="">>"">"
	else
		strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value="">"">"
    	strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value="">>"">"
	end if
    strDisplayList = strDisplayList & "<input type=""hidden"" name=""txtSearch"" value=" & strPalletSearch & ">"
	strDisplayList = strDisplayList & "<input type=""hidden"" name=""cboPalletType"" value=" & strPalletType & ">"
	strDisplayList = strDisplayList & "<br />"
    strDisplayList = strDisplayList & "Page: " & session("pallet_initial_page") & " to " & intpagecount
    strDisplayList = strDisplayList & "<br />"
	strDisplayList = strDisplayList & "<h2>Search results: " & intRecordCount & " pallets.</h2>"
    strDisplayList = strDisplayList & "</form>"
    strDisplayList = strDisplayList & "</td>"
    strDisplayList = strDisplayList & "</tr>"

    call CloseDataBase()
end sub

sub main	
	dim strPalletDepartment
	dim strPalletInfo
	dim intPalletStatus
				
	strPalletDepartment	= Request.Form("cboPalletDepartment")
	strPalletInfo 		= Trim(Request.Form("txtPalletInfo"))
	intPalletStatus		= Request.Form("cboPalletStatus")
		
	dim intPalletID
	dim strUpdateDepartment
	dim strUpdateInfo
	dim intUpdateStatus
	
	intPalletID			= Request("pallet_id")
	strUpdateDepartment	= Request.Form("txtUpdateDepartment")
	strUpdateInfo 		= Trim(Request.Form("txtUpdateInfo"))
	intUpdateStatus		= Request.Form("cboUpdateStatus")
	
	call UTL_validateLogin
	call setSearch

    if trim(session("pallet_initial_page"))  = "" then
    	session("pallet_initial_page") = 1
	end if    
	
	if Request.ServerVariables("REQUEST_METHOD") = "POST" then
		Select Case Trim(Request("action"))
			case "add"
				call addPallet(strPalletDepartment,strPalletInfo)
				'call displayPallets
			case "update"
				call updatePallet(intPalletID, strUpdateDepartment, strUpdateInfo, intUpdateStatus)
				'call displayPallets
		end select
	end if
	
	call displayPallets
end sub

call main

dim rs
dim intPageCount, intpage, intRecord
dim strDisplayList
dim strMessageText
%>
<table cellspacing="0" cellpadding="0" align="center" class="full_size_table" border="0">
  <!-- #include file="include/header.asp" -->
  <tr>
    <td class="first_content"><table cellpadding="5" cellspacing="0" border="0">
        <tr>
          <td valign="top"><img src="images/icon_pallet.jpg" border="0" alt="Pallet" /></td>
          <td valign="top"><div class="alert alert-search">
              <form name="frmSearch" id="frmSearch" action="list_pallet.asp?type=search" method="post" onsubmit="searchPallets()">
                <h3>Search Parameters:</h3>
                Pallet no:
                <input type="text" name="txtSearch" size="25" value="<%= request("txtSearch") %>" maxlength="20" />
                <select name="cboDepartment" onchange="searchPallets()">
                  <option <% if session("pallet_department") = "" then Response.Write " selected" end if%> value="">All Dept</option>
                  <option <% if session("pallet_department") = "AV" then Response.Write " selected" end if%> value="AV">AV</option>
                  <option <% if session("pallet_department") = "MPD" then Response.Write " selected" end if%> value="MPD">MPD</option>                 
                </select>
                <select name="cboStatus" onchange="searchPallets()">
                  <option <% if session("pallet_search_status") = "1" then Response.Write " selected" end if%> value="1">Open</option>
                  <option <% if session("pallet_search_status") = "0" then Response.Write " selected" end if%> value="0">Completed</option>
                </select>
                <input type="button" name="btnSearch" value="Search" onclick="searchPallets()" />
                <input type="button" name="btnReset" value="Reset" onclick="resetSearch()" />
              </form>
            </div></td>
          <td valign="top"><font color="green"><%= strMessageText %></font></td>
        </tr>
      </table>
      <p><a href="list_gra.asp">Goods Return BASE</a> &nbsp;-&nbsp; <a href="list_gra_report.asp">Report Summaries</a> &nbsp;-&nbsp; <a href="list_gra_report_writeoffs.asp">Write Offs Report</a> &nbsp;-&nbsp; <a href="list_gra_report_exported.asp">Exported Report</a> &nbsp;-&nbsp; <span class="current_header">Pallets</span></p>
      <div align="left">
      <form action="" method="post" name="form_add_pallet" id="form_add_pallet" onsubmit="return validateAddPalletForm(this)">
        <table cellpadding="5" cellspacing="0" class="item_maintenance_box">
          <tr>
            <td colspan="3" class="item_maintenance_header">Add New Pallet</td>
          </tr>
          <tr>
            <td width="40%">Department<span class="mandatory">*</span>:<br />
              <select name="cboPalletDepartment">
                <option value="">...</option>
                <option value="AV">AV</option>
                <option value="MPD">MPD</option>
                <option value="Other">Other</option>
              </select></td>
            <td width="40%">Info:<br />
              <input type="text" id="txtPalletInfo" name="txtPalletInfo" maxlength="30" size="40" /></td>
            <td width="20%" valign="bottom"><input type="hidden" name="Action" />
              <input type="submit" value="Add Pallet" /></td>
          </tr>
        </table>
      </form>
      </div>
      <br />
      <table cellspacing="0" cellpadding="8" class="database_records_nowidth" width="1200" border="0">
      <thead>
        <tr>
          
          <td>Pallet no</td>
          <td>Dept</td>
          <td>Pallet info</td>
          <td>Status</td>
          <td>&nbsp;</td>
          <td>Created</td>
          <td>Last modified</td>
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