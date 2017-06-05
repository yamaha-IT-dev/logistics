<% response.cookies("current_URL_cookie_logistics") = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?" & Request.Querystring %>
<!--#include file="include/connection_it.asp " -->
<!--#include file="class/clsUser.asp " -->
<% strSection = "user" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Users</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="include/javascript.js"></script>
<script type="text/javascript" src="include/generic_form_validations.js"></script>
<script language="JavaScript" type="text/javascript">
function searchItem(){    
    var strSearch 		= document.forms[0].txtSearch.value;
	var strDepartment  	= document.forms[0].cboDepartment.value;
	var strStatus 		= document.forms[0].cboStatus.value;

    document.location.href = 'list_user.asp?type=search&txtSearch=' + strSearch + '&cboDepartment=' + strDepartment + '&cboStatus=' + strStatus;	
}
    
function resetSearch(){
	document.location.href = 'list_user.asp?type=reset';    
}  

function validateNewUser(theForm) {
	var reason = "";
	var blnSubmit = true;

	reason += validateEmptyField(theForm.txtUsername,"Username");
	reason += validateSpecialCharacters(theForm.txtUsername,"Username");
	
	reason += validateEmptyField(theForm.txtFirstname,"Firstname");
	reason += validateSpecialCharacters(theForm.txtFirstname,"Firstname");
	
	reason += validateEmptyField(theForm.txtLastname,"Lastname");
	reason += validateSpecialCharacters(theForm.txtLastname,"Lastname");
	
	reason += validateEmail(theForm.txtEmail);
	
	reason += validateEmptyField(theForm.cboDepartment,"Department");
	
  	if (reason != "") {
    	alert("Some fields need correction:\n" + reason);    	
		blnSubmit = false;
		return false;
  	}

	if (blnSubmit == true){
        theForm.Action.value = 'Add';
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
			session("user_search") 		= ""
			session("user_department") 	= ""
			session("user_status") 		= ""
			session("user_initial_page") = 1
		case "search"
			session("user_search") 		= trim(Request("txtSearch"))
			session("user_department") 	= request("cboDepartment")			
			session("user_status") 		= Trim(Request("cboStatus"))
			session("user_initial_page") = 1
	end select
end sub

sub displayUser	
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
	
	
	
	strSQL = "SELECT * FROM tbl_users "
	strSQL = strSQL & "	WHERE division LIKE '%" & session("user_department") & "%' "
	strSQL = strSQL & "		AND (username LIKE '%" & session("user_search") & "%' "
	strSQL = strSQL & "			OR firstname LIKE '%" & session("user_search") & "%' "
	strSQL = strSQL & "			OR lastname LIKE '%" & session("user_search") & "%' "
	strSQL = strSQL & "			OR email LIKE '%" & session("user_search") & "%') "
	strSQL = strSQL & "		AND status LIKE '%" & session("user_status") & "%' "
	strSQL = strSQL & "	ORDER BY username"
			
	'Response.Write strSQL & "<br>"
	
	rs.Open strSQL, conn
			
	intPageCount = rs.PageCount
	intRecordCount = rs.recordcount
	
	Select Case Request("Action")
	    case "<<"
		    intpage = 1
			session("user_initial_page") = intpage
	    case "<"
		    intpage = Request("intpage") - 1
			session("user_initial_page") = intpage
			
			if session("user_initial_page") < 1 then session("user_initial_page") = 1
	    case ">"
		    intpage = Request("intpage") + 1
			session("user_initial_page") = intpage
			
			if session("user_initial_page") > intPageCount then session("user_initial_page") = IntPageCount
	    Case ">>"
		    intpage = intPageCount
			session("user_initial_page") = intpage	    
    end select

    strDisplayList = ""
	
	if not DB_RecSetIsEmpty(rs) Then
	
	    rs.AbsolutePage = session("user_initial_page")  
	
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
				
			strDisplayList = strDisplayList & "<td align=""center"" nowrap><a href=""update_user.asp?id=" & rs("user_id") & """><img src=""images/icon_view.png"" border=""0""></a></td>"			
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("username") & ""
			if DateDiff("d",rs("date_created"), strTodayDate) = 0 then
				strDisplayList = strDisplayList & " <img src=""images/icon_new.gif"" border=""0"">"
			end if
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("firstname") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("lastname") & "</td>"	
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("email") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("role_id") & "</td>"
			'strDisplayList = strDisplayList & "<td align=""center"">"
			'Select Case	rs("role_id")
			'	case 1
			'		strDisplayList = strDisplayList & "1"
			'	case 2
			'		strDisplayList = strDisplayList & "2"
			'	case 0
			'		strDisplayList = strDisplayList & "other"
			'end select						
			'strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("division") & "</td>"
			if rs("status") = 1 then 
				strDisplayList = strDisplayList & "<td class=""green_text"">Active</td>"
			else
				strDisplayList = strDisplayList & "<td class=""red_text"">In-active</td>"
			end if					
			strDisplayList = strDisplayList & "</tr>"
			
			rs.movenext
			iRecordCount = iRecordCount + 1
			If rs.EOF Then Exit For 
		next

	else
        strDisplayList = "<tr class=""innerdoc""><td colspan=""8"" align=""center"">No user records found.</td></tr>"
	end if
	
	strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
	strDisplayList = strDisplayList & "<td colspan=""8"" class=""recordspaging"">"
	strDisplayList = strDisplayList & "<form name=""MovePage"" action=""list_user.asp"" method=""post"">"
    strDisplayList = strDisplayList & "<input type=""hidden"" name=""intpage"" value=" & session("user_initial_page") & ">"
	
	if session("user_initial_page") = 1 then
   		strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value=""<<"">"
    	strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value=""<"">"
	else 
		strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value=""<<"">"
    	strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value=""<"">"
	end if	
	if session("user_initial_page") = intpagecount or intRecordCount = 0 then
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
    strDisplayList = strDisplayList & "Page: " & session("user_initial_page") & " to " & intpagecount
	strDisplayList = strDisplayList & "<br />"
	strDisplayList = strDisplayList & "Search results: " & intRecordCount & " records."
    strDisplayList = strDisplayList & "</form>"
    strDisplayList = strDisplayList & "</td>"
    strDisplayList = strDisplayList & "</tr>"
	
    call CloseDataBase()
end sub

sub main
	call UTL_validateLogin	
	call setSearch

    if trim(session("user_initial_page"))  = "" then
    	session("user_initial_page") = 1
	end if		
    
    call displayUser
	
	if Request.ServerVariables("REQUEST_METHOD") = "POST" then
		if trim(request("Action")) = "Add" then
			dim strUsername
			dim strFirstname
			dim strLastname
			dim strEmail
			dim strDepartment	
			dim intAdmin
			dim intManagerID
				
			strUsername 	= Trim(Request.Form("txtUsername"))
			strFirstname 	= Replace(Request.Form("txtFirstname"),"'","''")
			strLastname 	= Replace(Request.Form("txtLastname"),"'","''")
			strEmail		= Trim(Request.Form("txtEmail"))
			strDepartment	= Request.Form("cboDepartment")
			intAdmin		= Request.Form("chkAdmin")
			
			call addUser(strUsername,strFirstname,strLastname,strEmail,strDepartment,intAdmin)
			call displayUser
		end if
	end if
end sub

call main

dim rs
dim intPageCount, intpage, intRecord
dim strDisplayList
dim strDealerResultList
dim strStateList
dim strSalesManagerList
%>
<table cellspacing="0" cellpadding="0" align="center" class="full_size_table" border="0">
  <!-- #include file="include/header.asp" -->
  <tr>
    <td class="first_content"><table cellpadding="5" cellspacing="0" border="0">
        <tr>
          <td valign="top"><img src="images/icon_item-maintenance.jpg" border="0" alt="Item Maintenance" /></td>          
          <td valign="top"><div class="alert alert-search">
              <form name="frmSearch" id="frmSearch" action="list_user.asp?type=search" method="post" onsubmit="searchItem()">
                <h3>Search Parameters:</h3>
                Username / Firstname / Lastname / Email:
                <input type="text" name="txtSearch" size="25" value="<%= request("txtSearch") %>" maxlength="20" />
                <select name="cboDepartment" onchange="searchItem()">
                  <option value="">All Depts</option>
                  <option <% if session("user_department") = "AV" then Response.Write " selected" end if%> value="AV">AV</option>
                  <option <% if session("user_department") = "MPD" then Response.Write " selected" end if%> value="MPD">MPD</option>
                </select>
                <select name="cboStatus" onchange="searchItem()">
                  <option <% if session("user_status") = "" then Response.Write " selected" end if%> value="">All Status</option>
                  <option <% if session("user_status") = "1" then Response.Write " selected" end if%> value="1">Active</option>
                  <option <% if session("user_status") = "0" then Response.Write " selected" end if%> value="0">In-active</option>
                </select>
                <input type="button" name="btnSearch" value="Search" onclick="searchItem()" />
                <input type="button" name="btnReset" value="Reset" onclick="resetSearch()" />
              </form>
            </div></td>
        </tr>
      </table>
      <table>
        <tr>
          <td valign="top"><table cellspacing="0" cellpadding="4" class="database_records_nowidth" width="900">
              <tr class="innerdoctitle" align="center">
                <td>&nbsp;</td>
                <td>Username</td>
                <td>Firstname</td>
                <td>Lastname</td>
                <td>Email</td>
                <td>Role</td>
                <td>Division</td>
                <td>Status</td>
              </tr>
              <%= strDisplayList %>
            </table></td>
          <td valign="top" style="padding-left:10px;"><form action="" method="post" name="form_add_user" id="form_add_user" onsubmit="return validateNewUser(this)">
              <font color="green"><%= strMessageText %></font>
              <table cellpadding="5" cellspacing="0" class="item_maintenance_box">
                <tr>
                  <td colspan="2" class="item_maintenance_header">New User</td>
                </tr>
                <tr>
                  <td width="50%">Username<span class="mandatory">*</span>:<br />
                    <input type="text" id="txtUsername" name="txtUsername" maxlength="25" size="25" /></td>
                  <td width="50%" valign="top" align="right"></td>
                </tr>
                <tr>
                  <td>Firstname<span class="mandatory">*</span>:<br />
                    <input type="text" id="txtFirstname" name="txtFirstname" maxlength="20" size="25" /></td>
                  <td>Lastname<span class="mandatory">*</span>:<br />
                    <input type="text" id="txtLastname" name="txtLastname" maxlength="20" size="25" /></td>
                </tr>
                <tr>
                  <td colspan="2">Email<span class="mandatory">*</span>:<br />
                    <input type="text" id="txtEmail" name="txtEmail" maxlength="55" size="60" /></td>
                </tr>
                <tr>
                  <td colspan="2">Department<span class="mandatory">*</span>:<br />
                    <select name="cboDepartment">
                      <option value="">...</option>
                      <option value="AV">AV</option>
                      <option value="CC">CC</option>
                      <option value="FIN">FINANCE</option>
                      <option value="IT">IT</option>
                      <option value="LOG">LOGISTICS</option>
                      <option value="OP">OPERATIONS</option>
                      <option value="SER">SERVICE</option>
                      <option value="MPD">MPD</option>
                      <option value="PRO">PRO</option>
                      <option value="TRAD">TRAD</option>
                      <option value="CA">CA</option>
                      <option value="YMEC">YMEC</option>
                    </select></td>
                </tr>
                <tr>
                  <td colspan="2"><input type="hidden" name="Action" />
                    <input type="submit" value="Add" /></td>
                </tr>
              </table>
            </form></td>
        </tr>
      </table></td>
  </tr>
</table>
</body>
</html>