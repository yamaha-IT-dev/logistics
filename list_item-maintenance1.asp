<% response.cookies("current_URL_cookie_logistics") = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?" & Request.Querystring %>
<!--#include file="include/connection_it.asp " -->
<!--#include file="class/clsEmcApproval.asp " -->
<% strSection = "item" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Item Maintenance</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="include/javascript.js"></script>
<script language="JavaScript" type="text/javascript">
function searchItem(){    
    var strSearch 		= document.forms[0].txtSearch.value;
	var strDepartment  	= document.forms[0].cboDepartment.value;
	var strStatus 		= document.forms[0].cboStatus.value;

    document.location.href = 'list_item-maintenance1.asp?type=search&txtSearch=' + strSearch + '&cboDepartment=' + strDepartment + '&cboStatus=' + strStatus;	
}
    
function resetSearch(){
	document.location.href = 'list_item-maintenance1.asp?type=reset';    
}  

function submitApproval(theForm) {
	var blnSubmit = true;

	if (blnSubmit == true){
		theForm.Action.value = 'EMC Approved';		
		return true;		
    }
}

function submitRejection(theForm) {
	var blnSubmit = true;

	if (blnSubmit == true){
		theForm.Action.value = 'EMC Rejected';		
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
			session("item_search") 		= ""
			session("item_department") 	= ""
			session("item_status") 		= ""
			session("item_initial_page") = 1
		case "search"
			session("item_search") 		= trim(Request("txtSearch"))
			session("item_department") 	= request("cboDepartment")			
			session("item_status") 		= Trim(Request("cboStatus"))
			session("item_initial_page") = 1
	end select
end sub

sub displayItem	
	dim iRecordCount
	iRecordCount = 0
    dim strSortBy
	dim strSortItem
    dim strSQL
	dim strPageResultNumber
	dim strRecordPerPage
	dim intRecordCount
	dim strTodayDate
	
	strTodayDate = FormatDateTime(Date())
	
    call OpenDataBase()
	
	set rs = Server.CreateObject("ADODB.recordset")
	
	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic
	rs.PageSize = 100
	
	if session("item_status") = "" then
		session("item_status") = "1"
	end if
	
	strSQL = "SELECT * FROM yma_item_maintenance "
	strSQL = strSQL & "	WHERE department LIKE '%" & session("item_department") & "%' "
	strSQL = strSQL & "		AND (base_code LIKE '%" & session("item_search") & "%' "
	strSQL = strSQL & "			OR item_name LIKE '%" & session("item_search") & "%' "
	strSQL = strSQL & "			OR model_name LIKE '%" & session("item_search") & "%' "
	strSQL = strSQL & "			OR description LIKE '%" & session("item_search") & "%' "
	strSQL = strSQL & "			OR gmc_code LIKE '%" & session("item_search") & "%' "
	strSQL = strSQL & "			OR created_by LIKE '%" & session("item_search") & "%') "
	strSQL = strSQL & "		AND status LIKE '%" & session("item_status") & "%' "
	strSQL = strSQL & "	ORDER BY date_created DESC"
			
	'Response.Write strSQL & "<br>"
	
	rs.Open strSQL, conn
			
	intPageCount = rs.PageCount
	intRecordCount = rs.recordcount
	
	Select Case Request("Action")
	    case "<<"
		    intpage = 1
			session("item_initial_page") = intpage
	    case "<"
		    intpage = Request("intpage") - 1
			session("item_initial_page") = intpage
			
			if session("item_initial_page") < 1 then session("item_initial_page") = 1
	    case ">"
		    intpage = Request("intpage") + 1
			session("item_initial_page") = intpage
			
			if session("item_initial_page") > intPageCount then session("item_initial_page") = IntPageCount
	    Case ">>"
		    intpage = intPageCount
			session("item_initial_page") = intpage
    end select

    strDisplayList = ""
	
	if not DB_RecSetIsEmpty(rs) Then
	
	    rs.AbsolutePage = session("item_initial_page")  
	
		For intRecord = 1 To rs.PageSize 
			if (DateDiff("d",rs("date_modified"), strTodayDate) = 0) then
				if iRecordCount Mod 2 = 0 then
					strDisplayList = strDisplayList & "<tr class=""updated_today"">"
				else
					strDisplayList = strDisplayList & "<tr class=""updated_today_2"">"
				end if
			else
				'strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
				if iRecordCount Mod 2 = 0 then
					strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
				else
					strDisplayList = strDisplayList & "<tr class=""innerdoc_2"">"
				end if
			end if
				
			strDisplayList = strDisplayList & "<td align=""center"" nowrap><a href=""update_item-maintenance.asp?id=" & rs("item_id") & """><img src=""images/icon_view.png"" border=""0""></a></td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("created_by") & " - " & WeekDayName(WeekDay(rs("date_created"))) & ", " & FormatDateTime(rs("date_created"),1) & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("department") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("base_code") & ""
			if DateDiff("d",rs("date_created"), strTodayDate) = 0 then
				strDisplayList = strDisplayList & " <img src=""images/icon_new.gif"" border=""0"">"
			end if
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("model_name") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("description") & "</td>"	
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("gmc_code") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">"			
			if rs("multicolour") = 1 then
				strDisplayList = strDisplayList & "Yes"
			else
				strDisplayList = strDisplayList & "-"
			end if			
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">$" & rs("rrp") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">"			

			Select Case	rs("set_item")
				case 1
					strDisplayList = strDisplayList & "Set"
				case 2
					strDisplayList = strDisplayList & "Kit"
				case 0
					strDisplayList = strDisplayList & "-"
			end select	
						
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">"			
			if rs("mod_required") = 1 then
				strDisplayList = strDisplayList & "Yes"
			else
				strDisplayList = strDisplayList & "-"
			end if			
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">"	
				
			Select Case	rs("marketing_approval")
				case 1
					strDisplayList = strDisplayList & "<img src=images/tick.gif>"
				case 0
					strDisplayList = strDisplayList & "<img src=images/cross.gif>"
				case else
					strDisplayList = strDisplayList & "-"
			end select	
				
			strDisplayList = strDisplayList & "</td>"
			
			strDisplayList = strDisplayList & "<td align=""center"">"		
			
			Select Case	rs("gm_approval")
				case 1
					strDisplayList = strDisplayList & "<img src=images/tick.gif>"
				case 0
					strDisplayList = strDisplayList & "<img src=images/cross.gif>"
				case else
					strDisplayList = strDisplayList & "Not yet"
			end select
		
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">"	
			
			Select Case	rs("emc_approval")
				case 1
					strDisplayList = strDisplayList & "<img src=images/tick.gif>"
				case 0
					strDisplayList = strDisplayList & "<img src=images/cross.gif>"
				case else
					'strDisplayList = strDisplayList & "Not yet"
					strDisplayList = strDisplayList & "<table><tr><td align=""center"">"
					strDisplayList = strDisplayList & "		<form method=""post"" name=""form_approve"" id=""form_approve"" onsubmit=""return submitApproval(this)"">"
					strDisplayList = strDisplayList & "			<input type=""hidden"" name=""action"" value=""EMC Approved"">"
					strDisplayList = strDisplayList & "			<input type=""hidden"" name=""item_id"" value=""" & rs("item_id") & """>"
					strDisplayList = strDisplayList & "			<input type=""hidden"" name=""base_code"" value=""" & rs("base_code") & """>"
					strDisplayList = strDisplayList & "			<input type=""submit"" value=""Approve"" style=""color:green"" />"
					strDisplayList = strDisplayList & "		</form>"
					strDisplayList = strDisplayList & "	</td>"
					strDisplayList = strDisplayList & "	<td align=""center"">"
					strDisplayList = strDisplayList & "		<form method=""post"" name=""form_reject"" id=""form_reject"" onsubmit=""return submitRejection(this)"">"
					strDisplayList = strDisplayList & "			<input type=""hidden"" name=""action"" value=""EMC Rejected"">"
					strDisplayList = strDisplayList & "			<input type=""hidden"" name=""item_id"" value=""" & rs("item_id") & """>"
					strDisplayList = strDisplayList & "			<input type=""hidden"" name=""base_code"" value=""" & rs("base_code") & """>"
					strDisplayList = strDisplayList & "			<input type=""submit"" value=""Reject"" style=""color:red"" />"
					strDisplayList = strDisplayList & "		</form>"
					strDisplayList = strDisplayList & "	</td></tr></table>"
			end select		
		
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">"	
				
			Select Case	rs("logistics_processed")
				case 1
					strDisplayList = strDisplayList & "<img src=images/tick.gif>"
				case 0
					strDisplayList = strDisplayList & "<img src=images/cross.gif>"
				case else
					strDisplayList = strDisplayList & "Not yet"
			end select	
				
			strDisplayList = strDisplayList & "</td>"
			if rs("status") = 1 then 
				strDisplayList = strDisplayList & "<td align=""center"">Open</td>"
			else
				strDisplayList = strDisplayList & "<td class=""green_text"">Completed</td>"
			end if
			
			'strDisplayList = strDisplayList & "<td align=""center""><a onclick=""return confirm('Are you sure you want to delete " & rs("base_code") & " ?');"" href='delete_item-maintenance.asp?id=" & rs("item_id") & "'><img src=""images/btn_delete.gif"" border=""0""></a></td>"			
			strDisplayList = strDisplayList & "</tr>"
			
			rs.movenext
			iRecordCount = iRecordCount + 1
			If rs.EOF Then Exit For 
		next

	else
        strDisplayList = "<tr class=""innerdoc""><td colspan=""20"" align=""center"">No items found.</td></tr>"
	end if
	
	strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
	strDisplayList = strDisplayList & "<td colspan=""20"" class=""recordspaging"">"
	strDisplayList = strDisplayList & "<form name=""MovePage"" action=""list_item-maintenance1.asp"" method=""post"">"
    strDisplayList = strDisplayList & "<input type=""hidden"" name=""intpage"" value=" & session("item_initial_page") & ">"
	
	if session("item_initial_page") = 1 then
   		strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value=""<<"">"
    	strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value=""<"">"
	else 
		strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value=""<<"">"
    	strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value=""<"">"
	end if	
	if session("item_initial_page") = intpagecount or intRecordCount = 0 then
    	strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value="">"">"
    	strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value="">>"">"
	else
		strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value="">"">"
    	strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value="">>"">"
	end if
    strDisplayList = strDisplayList & "<input type=""hidden"" name=""txtSearch"" value=" & strSearch & ">"
	strDisplayList = strDisplayList & "<input type=""hidden"" name=""cboDepartment"" value=" & strItemDepartment & ">"
	strDisplayList = strDisplayList & "<input type=""hidden"" name=""cboStatus"" value=" & strStatus & ">"
    strDisplayList = strDisplayList & "<br />"
    strDisplayList = strDisplayList & "Page: " & session("item_initial_page") & " to " & intpagecount
	strDisplayList = strDisplayList & "<br />"
	strDisplayList = strDisplayList & "<h2>Search results: " & intRecordCount & " items.</h2>"
    strDisplayList = strDisplayList & "</form>"
    strDisplayList = strDisplayList & "</td>"
    strDisplayList = strDisplayList & "</tr>"
	
    call CloseDataBase()
end sub

sub main
	call UTL_validateLogin
	select case trim(session("UsrUserName"))
		case "simong"
			session("item_department") = "AV"
		case "justind"
			session("item_department") = "AV"
		case "dalem"
			session("item_department") = "AV"
		case "marka"
			session("item_department") = "MPD"
		case else
			session("item_department") = request("cboDepartment")
	end select
	
	call setSearch

    if trim(session("item_initial_page"))  = "" then
    	session("item_initial_page") = 1
	end if
	
	if Request.ServerVariables("REQUEST_METHOD") = "POST" then
		intItemID 	= Trim(Request("item_id"))
		strBaseCode = Trim(Request("base_code"))
		
		select case Trim(Request("Action"))			
			case "EMC Approved"
				call approveEMC(intItemID, strBaseCode, session("UsrUserName"))
			case "EMC Rejected"
				call rejectEMC(intItemID, strBaseCode, session("UsrUserName"))			
		end select
	end if
	
	call displayItem 
end sub

call main

dim rs
dim intPageCount, intpage, intRecord
dim strDisplayList
dim strDealerResultList
dim strStateList
dim strSalesManagerList

Dim intItemID
Dim strBaseCode
%>
<table cellspacing="0" cellpadding="0" align="center" class="full_size_table" border="0">
  <!-- #include file="include/header.asp" -->
  <tr>
    <td class="first_content"><table cellpadding="5" cellspacing="0" border="0">
        <tr>
          <td valign="top"><img src="images/icon_item-maintenance.jpg" border="0" alt="Item Maintenance" /></td>
          <td valign="top"><div class="alert alert-success"><img src="images/add_icon.png" border="0" align="bottom" /> <a href="add_item-maintenance.asp">Add Item Maintenance</a></div>
            <p><img src="images/icon_document.png" border="0" align="bottom" /> <a href="How-to-Item-Maintenance.docx" target="_blank">Download the instruction</a></p></td>
          <td valign="top"><div class="alert alert-search">
              <form name="frmSearch" id="frmSearch" action="list_item-maintenance1.asp?type=search" method="post" onsubmit="searchItem()">
                <h3>Search Parameters:</h3>
                Created by / BASE code / Model name / Description / GMC code:
                <input type="text" name="txtSearch" size="25" value="<%= request("txtSearch") %>" maxlength="20" />
                <select name="cboDepartment" onchange="searchItem()">
                  <option value="">All Depts</option>
                  <option <% if session("item_department") = "AV" then Response.Write " selected" end if%> value="AV">AV</option>
                  <option <% if session("item_department") = "MPD" then Response.Write " selected" end if%> value="MPD">MPD</option>
                </select>
                <select name="cboStatus" onchange="searchItem()">
                  <option <% if session("item_status") = "1" then Response.Write " selected" end if%> value="1">Open</option>
                  <option <% if session("item_status") = "0" then Response.Write " selected" end if%> value="0">Completed</option>
                </select>
                <input type="button" name="btnSearch" value="Search" onclick="searchItem()" />
                <input type="button" name="btnReset" value="Reset" onclick="resetSearch()" />
              </form>
            </div></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td><table cellspacing="0" cellpadding="4" class="database_records">
        <tr class="innerdoctitle" align="center">
          <td>&nbsp;</td>
          <td>Created by</td>
          <td>Dept</td>
          <td>BASE code</td>
          <td>Model name</td>
          <td>Description</td>
          <td>GMC</td>
          <td>Multicolour</td>
          <td>RRP</td>
          <td>Set / Kit</td>
          <td>Mod required</td>
          <td>AV marketing approval</td>
          <td>GM approval</td>
          <td>EMC approval</td>
          <td>Logistics processed</td>
          <td>Status</td>
        </tr>
        <%= strDisplayList %>
      </table></td>
  </tr>
</table>
</body>
</html>