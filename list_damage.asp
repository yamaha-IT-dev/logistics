<% response.cookies("current_URL_cookie_logistics") = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?" & Request.Querystring %>
<!--#include file="include/connection_it.asp " -->
<!--#include file="class/clsDamageType.asp" -->
<% strSection = "damage" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Warehouse Damage</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="include/javascript.js"></script>
<script type="text/javascript" src="include/generic_form_validations.js"></script>
<script language="JavaScript" type="text/javascript">
function searchDamages(){
    var strDamageSearch = document.forms[0].txtSearch.value;
	var strDamageType  	= document.forms[0].cboDamageType.value;
	var strDamageSort  	= document.forms[0].cboSort.value;
	var strDamageYear  	= document.forms[0].cboYear.value;
	var strStatus 		= document.forms[0].cboStatus.value;

    document.location.href = 'list_damage.asp?type=search&txtSearch=' + strDamageSearch + '&cboDamageType=' + strDamageType + '&cboSort=' + strDamageSort + '&cboYear=' + strDamageYear + '&cboStatus=' + strStatus;
}

function resetSearch(){
	document.location.href = 'list_damage.asp?type=reset';
}

function validateFormOnSubmit(theForm) {
	var reason = "";
	var blnSubmit = true;

	reason += validateSpecialCharacters(theForm.txtLocation,"Location");
	reason += validateSpecialCharacters(theForm.txtConnote,"Connote");
	
  	if (reason != "") {
    	alert("Some fields need correction:\n" + reason);

		blnSubmit = false;
		return false;
  	}

	if (blnSubmit == true){
        theForm.Action.value = 'Update';
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
			session("damage_search") 		= ""
			session("damage_search_type") 	= ""
			session("damage_sort") 			= ""
			session("damage_year") 			= ""
			session("damage_search_status") = ""
			session("damage_initial_page") 	= 1
		case "search"
			session("damage_search") 		= Trim(Request("txtSearch"))
			session("damage_search_type") 	= Trim(Request("cboDamageType"))
			session("damage_sort") 			= Trim(Request("cboSort"))
			session("damage_year") 			= Trim(Request("cboYear"))
			session("damage_search_status") = Trim(Request("cboStatus"))
			session("damage_initial_page") 	= 1
	end select
end sub

sub displayDamagedItems
	dim iRecordCount
	iRecordCount = 0
    dim strDamageSortBy
	dim strDamageSortItem
    dim strDamageSearch
    dim strSQL
	'dim strDamageType
	dim strDamageSort
	dim strStatus

	dim strPageResultNumber
	dim strRecordPerPage
	dim intRecordCount
	dim strModifiedDate

	dim strTodayDate
	strTodayDate = FormatDateTime(Date())

	if session("damage_search_status") = "" then
		session("damage_search_status") = "1"
	end if

	
	'	session("damage_year") = "2013"
	'end if

	if session("damage_sort") = "" then
		session("damage_sort") = "date_created DESC"
	end if

    call OpenDataBase()

	set rs = Server.CreateObject("ADODB.recordset")

	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic
	rs.PageSize = 100

	strSQL = "SELECT * FROM yma_damage "
	strSQL = strSQL & "	WHERE "
	if session("damage_year") <> "" then
		strSQL = strSQL & " 	YEAR(date_created) = '" & trim(session("damage_year")) & "' AND "
	end if
	strSQL = strSQL & "		damage_type LIKE '%" & session("damage_search_type") & "%' "
	strSQL = strSQL & "		AND (damage_item LIKE '%" & session("damage_search") & "%' "
	strSQL = strSQL & "			OR damage_serial_no LIKE '%" & session("damage_search") & "%') "
	strSQL = strSQL & "		AND status LIKE '%" & session("damage_search_status") & "%' "
	strSQL = strSQL & "	ORDER BY " & session("damage_sort")

	'Response.Write strSQL

	rs.Open strSQL, conn

	intPageCount = rs.PageCount
	intRecordCount = rs.recordcount
	
	Select Case Request("Action")
	    case "<<"
		    intpage = 1
			session("damage_initial_page") = intpage
	    case "<"
		    intpage = Request("intpage") - 1
			session("damage_initial_page") = intpage

			if session("damage_initial_page") < 1 then session("damage_initial_page") = 1
	    case ">"
		    intpage = Request("intpage") + 1
			session("damage_initial_page") = intpage

			if session("damage_initial_page") > intPageCount then session("damage_initial_page") = IntPageCount
	    Case ">>"
		    intpage = intPageCount
			session("damage_initial_page") = intpage
    end select

    strDisplayList = ""

	if not DB_RecSetIsEmpty(rs) Then

		For intRecord = 1 To rs.PageSize
		
			strDisplayList = strDisplayList & "<form method=""post"" name=""form_update_damage"" id=""form_update_damage"" onsubmit=""return validateFormOnSubmit(this)"">"
			strDisplayList = strDisplayList & "<input type=""hidden"" name=""action"" value=""update"">"
			strDisplayList = strDisplayList & "<input type=""hidden"" name=""id"" value=""" & trim(rs("damage_id")) & """>"
			
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

			strDisplayList = strDisplayList & "<td><a href=""update_damage.asp?id=" & rs("damage_id") & """>" & rs("damage_id") & "</a></td>"
			strDisplayList = strDisplayList & "<td>" & rs("damage_item") & ""
			if DateDiff("d",rs("date_created"), strTodayDate) = 0 then
				strDisplayList = strDisplayList & " <img src=""images/icon_new.gif"" border=""0"">"
			end if
			strDisplayList = strDisplayList & "</td>"

			if Session("UsrLoginRole") = 1 then
				strDisplayList = strDisplayList & "<td><i>$" & rs("lic") & "</i></td>"
			end if

			strDisplayList = strDisplayList & "<td>" & rs("damage_serial_no") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("damage_type") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("course_damage") & "</td>"

			strDisplayList = strDisplayList & "<td>"
			if rs("sent_excel") = 1 then
				strDisplayList = strDisplayList & "<img src=images/tick.gif>"
			else
				strDisplayList = strDisplayList & "<img src=images/cross.gif>"
			end if
			strDisplayList = strDisplayList & "</td>"

			if rs("sent_excel_date") = "01/01/1900" or rs("sent_excel_date") = "1/1/1900" then
				strDisplayList = strDisplayList & "<td class=""orange_text"">NA</td>"
			else
				strDisplayList = strDisplayList & "<td>" & WeekDayName(WeekDay(rs("sent_excel_date"))) & ", " & FormatDateTime(rs("sent_excel_date"),1) & "</td>"
			end if
			'strDisplayList = strDisplayList & "<td><input type=""text"" id=""txtSentDate"" name=""txtSentDate"" maxlength=""12"" size=""15"" value=""" & rs("sent_excel_date") & """ ></td>"
			strDisplayList = strDisplayList & "<td><input type=""text"" id=""txtLocation"" name=""txtLocation"" maxlength=""12"" size=""15"" value=""" & rs("damage_location") & """ ></td>"
			strDisplayList = strDisplayList & "<td><input type=""text"" id=""txtConnote"" name=""txtConnote"" maxlength=""12"" size=""15"" value=""" & rs("damage_connote") & """ ></td>"
			strDisplayList = strDisplayList & "<td><input type=""submit"" value=""Update"" /></td>"	
			strDisplayList = strDisplayList & "<td>" & rs("damage_comments") & "</td>"
			if rs("status") = 1 then
				strDisplayList = strDisplayList & "<td>Open</td>"
			else
				strDisplayList = strDisplayList & "<td class=""green_text"">Completed</td>"
			end if

			strDisplayList = strDisplayList & "<td><a onclick=""return confirm('Are you sure you want to delete " & rs("damage_item") & " ?');"" href='delete_damage.asp?damage_id=" & rs("damage_id") & "'><img src=""images/btn_delete.gif"" border=""0""></a></td>"
			strDisplayList = strDisplayList & "</tr>"
			strDisplayList = strDisplayList & "</form>"
			rs.movenext
			iRecordCount = iRecordCount + 1
			If rs.EOF Then Exit For
		next

	else
        strDisplayList = "<tr class=""innerdoc""><td colspan=""14"" align=""center"">No damaged items found.</td></tr>"
	end if

	strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
	strDisplayList = strDisplayList & "<td colspan=""14"" class=""recordspaging"">"
	strDisplayList = strDisplayList & "<form name=""MovePage"" action=""list_damage.asp"" method=""post"">"
    strDisplayList = strDisplayList & "<input type=""hidden"" name=""intpage"" value=" & session("damage_initial_page") & ">"

	if session("damage_initial_page") = 1 then
   		strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value=""<<"">"
    	strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value=""<"">"
	else
		strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value=""<<"">"
    	strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value=""<"">"
	end if
	if session("damage_initial_page") = intpagecount or intRecordCount = 0 then
    	strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value="">"">"
    	strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value="">>"">"
	else
		strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value="">"">"
    	strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value="">>"">"
	end if
    strDisplayList = strDisplayList & "<input type=""hidden"" name=""txtSearch"" value=" & strDamageSearch & ">"
	strDisplayList = strDisplayList & "<input type=""hidden"" name=""cboDamageType"" value=" & strDamageType & ">"
	strDisplayList = strDisplayList & "<input type=""hidden"" name=""cboSort"" value=" & strDamageSort & ">"
	strDisplayList = strDisplayList & "<br />"
    strDisplayList = strDisplayList & "Page: " & session("damage_initial_page") & " to " & intpagecount
    strDisplayList = strDisplayList & "<br />"
	strDisplayList = strDisplayList & "<h2>Search results: " & intRecordCount & " damaged items.</h2>"
    strDisplayList = strDisplayList & "</form>"
    strDisplayList = strDisplayList & "</td>"
    strDisplayList = strDisplayList & "</tr>"

    call CloseDataBase()
end sub

sub updateDamage
	dim strSQL
	dim intID		
	dim strLocation
	dim strConnote
	
	intID 			= Request.Form("id")
	strLocation 	= Trim(Request.Form("txtLocation"))
	strConnote 		= Trim(Request.Form("txtConnote"))
	
	Call OpenDataBase()

	strSQL = "UPDATE yma_damage SET "
	strSQL = strSQL & "damage_location = '" & Server.HTMLEncode(strLocation) & "',"
	strSQL = strSQL & "damage_connote = '" & Server.HTMLEncode(strConnote) & "',"
	strSQL = strSQL & "date_modified = getdate(),"
	strSQL = strSQL & "modified_by = '" & session("UsrUserName") & "' WHERE damage_id = " & intID

	'response.Write strSQL
	on error resume next
	conn.Execute strSQL

	On error Goto 0

	if err <> 0 then
		strMessageText = err.description
	else
		strMessageText = "The record has been updated."
	end if

	Call CloseDataBase()
end sub

sub main
	call UTL_validateLogin
	call setSearch

    if trim(session("damage_initial_page"))  = "" then
    	session("damage_initial_page") = 1
	end if

    call displayDamagedItems
	call getDamageTypeList
	
	if Request.ServerVariables("REQUEST_METHOD") = "POST" then	
		Select Case Trim(Request("action"))
			case "update"			
				call updateDamage
				call displayDamagedItems		
		end select
	end if
end sub

call main

dim rs
dim intPageCount, intpage, intRecord
dim strDisplayList
dim strDealerResultList
dim strDamageTypeList
dim strMessageText

%>
<table cellspacing="0" cellpadding="0" align="center" class="full_size_table" border="0">
  <!-- #include file="include/header.asp" -->
  <tr>
    <td class="first_content"><table cellpadding="5" cellspacing="0" border="0">
        <tr>
          <td valign="top"><img src="images/icon_damaged.jpg" border="0" alt="Damage Stocks" /></td>
          <td valign="top"><div class="alert alert-success"><img src="images/add_icon.png" border="0" align="bottom" /> <a href="add_damage.asp">Add Warehouse Damage</a></div>
            <% if Session("UsrLoginRole") = 1 then %>
            <p><img src="images/icon_excel.jpg" border="0" /> <a href="export_damages.asp?search=<%= session("damage_search") %>&type=<%= session("damage_search_type") %>&year=<%= session("damage_year") %>&status=<%= session("damage_search_status") %>&sort=<%= session("damage_sort") %>">Export</a></p>
            <% end if %></td>
          <td valign="top"><div class="alert alert-search">
              <form name="frmSearch" id="frmSearch" action="list_damage.asp?type=search" method="post" onsubmit="searchDamages()">
                <h3>Search Parameters:</h3>
                Item / Serial no:
                <input type="text" name="txtSearch" size="25" value="<%= request("txtSearch") %>" maxlength="20" />
                <select name="cboDamageType" onchange="searchDamages()">
                  <option value=''>All Damage Types</option>
                  <%= strDamageTypeList %>
                </select>
                <select name="cboYear" onchange="searchDamages()">
                  <option <% if session("damage_year") = "" then Response.Write " selected" end if%> value="">All years</option>
				  <option <% if session("damage_year") = "2016" then Response.Write " selected" end if%> value="2016">2016 only</option>
                  <option <% if session("damage_year") = "2015" then Response.Write " selected" end if%> value="2015">2015 only</option>
                  <option <% if session("damage_year") = "2014" then Response.Write " selected" end if%> value="2014">2014 only</option>
                  <option <% if session("damage_year") = "2013" then Response.Write " selected" end if%> value="2013">2013 only</option>
                  <option <% if session("damage_year") = "2012" then Response.Write " selected" end if%> value="2012">2012 only</option>
                </select>
                <select name="cboStatus" onchange="searchDamages()">
                  <option <% if session("damage_search_status") = "1" then Response.Write " selected" end if%> value="1">Open</option>
                  <option <% if session("damage_search_status") = "0" then Response.Write " selected" end if%> value="0">Completed</option>
                </select>
                <select name="cboSort" onchange="searchDamages()">
                  <option value="">Sort by...</option>
                  <option <% if session("damage_sort") = "damage_item" then Response.Write " selected" end if%> value="damage_item">Damage Item</option>
                  <option <% if session("damage_sort") = "damage_serial_no" then Response.Write " selected" end if%> value="damage_serial_no">Serial No</option>
                  <option <% if session("damage_sort") = "damage_type" then Response.Write " selected" end if%> value="damage_type">Damage Type</option>
                  <option <% if session("damage_sort") = "course_damage" then Response.Write " selected" end if%> value="course_damage">Course of Damage</option>
                  <option <% if session("damage_sort") = "date_created" then Response.Write " selected" end if%> value="date_created">Date Created</option>
                </select>
                <input type="button" name="btnSearch" value="Search" onclick="searchDamages()" />
                <input type="button" name="btnReset" value="Reset" onclick="resetSearch()" />
              </form>
            </div></td>
        </tr>
      </table>
      <p><font color="green"><%= strMessageText %></font></p></td>
  </tr>
  <tr>
    <td><table cellspacing="0" cellpadding="8" class="database_records">
    <thead>
        <tr>
          <td width="4%">ID</td>
          <td width="10%">Item</td>
          <% if Session("UsrLoginRole") = 1 then %>
          <td width="2%">LIC</td>
          <% end if %>
          <td width="10%">Serial no</td>
          <td width="4%">Type</td>
          <td width="10%">Cause</td>
          <td width="5%">Sent?</td>
          <td width="12%">Sent to Excel Date</td>
          <td width="10%">Location</td>
          <td width="10%">Connote</td>
          <td width="3%"></td>
          <td width="23%">Comments</td>
          <td width="5%">Status</td>
          <td width="2%">&nbsp;</td>
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