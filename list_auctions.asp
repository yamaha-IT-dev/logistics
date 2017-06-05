<% response.cookies("current_URL_cookie_logistics") = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?" & Request.Querystring %>
<!--#include file="include/connection_it.asp " -->
<% strSection = "auction" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Auction Winners - 15 Aug '12</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="include/javascript.js"></script>
<script type="text/javascript" src="include/generic_form_validations.js"></script>
<script language="JavaScript" type="text/javascript">
function searchAuction(){
    var strSearch = document.forms[0].txtSearch.value;

     document.location.href = 'list_auctions.asp?type=search&txtSearch=' + strSearch;
}

function resetSearch(){
	document.location.href = 'list_auctions.asp?type=reset';
}

function validateFormOnSubmit(theForm) {
	var reason = "";
	var blnSubmit = true;

	reason += validateEmptyField(theForm.txtSalesOrderNo,"Sales order no");
	reason += validateSpecialCharacters(theForm.txtSalesOrderNo,"Sales order no");
	reason += validateSpecialCharacters(theForm.txtInvoiceNo,"Invoice no");
	//reason += validateSpecialCharacters(theForm.txtComments,"Comments");
	
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
			session("auction_search") = ""
		case "search"
			session("auction_search") = Trim(Request("txtSearch"))
	end select
end sub

sub displayAuction
	dim iRecordCount
	iRecordCount = 0
    Dim strSearch
    dim strSQL
	dim strType

	dim strPageResultNumber
	dim strRecordPerPage
	dim intRecordCount
	dim strModifiedDate

	dim strTodayDate
	strTodayDate = FormatDateTime(Date())

	session("auction_search") = Trim(Request("txtSearch"))
	'Response.Write "session search: " & session("auction_search")
	strSearch = ""
    if len(trim(Session("auction_search"))) > 0 then
        strSearch = Session("auction_search")
    end if

    call OpenDataBase()

	set rs = Server.CreateObject("ADODB.recordset")

	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic

	strRecordPerPage = 900

	rs.PageSize = 900

	strSQL = "SELECT * FROM yma_auction_aug2012 "
	strSQL = strSQL & " WHERE (product LIKE '%" & strSearch & "%' "
	strSQL = strSQL & "			OR winner LIKE '%" & strSearch & "%' "
	strSQL = strSQL & "			OR category LIKE '%" & strSearch & "%')"

	rs.Open strSQL, conn

	'Response.Write strSQL

	intPageCount = rs.PageCount
	intRecordCount = rs.recordcount

    strDisplayList = ""

	if not DB_RecSetIsEmpty(rs) Then

	    'rs.AbsolutePage = session("cinitialPage")

		For intRecord = 1 To rs.PageSize
		
			strDisplayList = strDisplayList & "<form method=""post"" name=""form_update_action"" id=""form_update_action"" onsubmit=""return validateFormOnSubmit(this)"">"
			strDisplayList = strDisplayList & "<input type=""hidden"" name=""action"" value=""update"">"
			strDisplayList = strDisplayList & "<input type=""hidden"" name=""id"" value=""" & trim(rs("id")) & """>"
		
			if (DateDiff("d",rs("date_modified"), strTodayDate) = 0) then
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

			strDisplayList = strDisplayList & "<td align=""center""><a href=""update_auction.asp?id=" & rs("id") & """>Edit</a></td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("id") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("category") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("product") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("basecode") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("comments") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">$" & rs("lic") & "</td>"
			
			strDisplayList = strDisplayList & "<td align=""center"">$" & rs("reserve") & "</td>"
			if rs("winning_bid") <> "" then
				strDisplayList = strDisplayList & "<td align=""center"">$" & rs("winning_bid") & "</td>"
			else
				strDisplayList = strDisplayList & "<td align=""center"">-</td>"
			end if
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("winner") & "</td>"
			
			strDisplayList = strDisplayList & "<td align=""center""><input type=""text"" id=""txtSalesOrderNo"" name=""txtSalesOrderNo"" maxlength=""6"" size=""8"" value=""" & rs("sales_order_no") & """ ></td>"
			
			strDisplayList = strDisplayList & "<td align=""center""><input type=""text"" id=""txtInvoiceNo"" name=""txtInvoiceNo"" maxlength=""7"" size=""8"" value=""" & rs("invoice_no") & """ ></td>"
			
			'strDisplayList = strDisplayList & "<td align=""center""><input type=""text"" id=""txtComments"" name=""txtComments"" maxlength=""80"" size=""30"" value=""" & rs("comments") & """ ></td>"
			strDisplayList = strDisplayList & "<td align=""center""><input type=""submit"" value=""Update"" /></td>"
			strDisplayList = strDisplayList & "</tr>"
			strDisplayList = strDisplayList & "</form>"
			rs.movenext
			iRecordCount = iRecordCount + 1
			If rs.EOF Then Exit For
		next
	else
        strDisplayList = "<tr class=""innerdoc""><td colspan=""20"" align=""center"">There is no auction items.</td></tr>"
	end if

	strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
	strDisplayList = strDisplayList & "<td colspan=""20"" class=""recordspaging"">"
    strDisplayList = strDisplayList & "<input type=""hidden"" name=""txtSearch"" value=" & strSearch & ">"
    strDisplayList = strDisplayList & "<br />"
	strDisplayList = strDisplayList & "Search results: " & intRecordCount & " records."
    'strDisplayList = strDisplayList & "</form>"
    strDisplayList = strDisplayList & "</td>"
    strDisplayList = strDisplayList & "</tr>"

    call CloseDataBase()
end sub

sub updateAuction
	dim strSQL
	dim intID		
	dim strSalesOrderNo
	dim strInvoiceNo
	dim strComments
	
	intID 			= Request.Form("id")
	strSalesOrderNo = Trim(Request.Form("txtSalesOrderNo"))
	strInvoiceNo 	= Trim(Request.Form("txtInvoiceNo"))
	strComments 	= Trim(Request.Form("txtComments"))
	
	Call OpenDataBase()

	strSQL = "UPDATE yma_auction_aug2012 SET "
	strSQL = strSQL & "sales_order_no = '" & Server.HTMLEncode(strSalesOrderNo) & "',"
	strSQL = strSQL & "invoice_no = '" & Server.HTMLEncode(strInvoiceNo) & "',"
	'strSQL = strSQL & "comments = '" & Server.HTMLEncode(strComments) & "',"
	strSQL = strSQL & "date_modified = getdate(),"
	strSQL = strSQL & "modified_by = '" & session("UsrUserName") & "' WHERE id = " & intID

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

    if trim(session("cinitialPage"))  = "" then
    	session("cinitialPage") = 1
	end if

    call displayAuction
	if Request.ServerVariables("REQUEST_METHOD") = "POST" then	
		Select Case Trim(Request("action"))
			case "update"			
				call updateAuction
				call displayAuction		
		end select
	end if
end sub

call main

dim strMessageText
dim rs
dim intPageCount, intpage, intRecord
dim strDisplayList
%>
<table cellspacing="0" cellpadding="0" align="center" class="full_size_table">
  <!-- #include file="include/header.asp" -->
  <tr>
    <td class="first_content"><table cellpadding="5" cellspacing="0" border="0">
        <tr>
          <td valign="top"><img src="../auction/images/ybay_logo_clean.jpg" border="0" alt="yBay" /></td>
          <td valign="top"><div class="alert alert-search">
              <form name="frmSearch" id="frmSearch">
                <h3>Search parameters</h3>
                Product / Category / Winner:
                <input type="text" name="txtSearch" size="15" value="<%= request("txtSearch") %>" />
                <input type="button" name="btnSearch" value="Search" onclick="searchAuction()" />
                <input type="button" name="btnReset" value="Reset" onclick="resetSearch()" />
              </form>
            </div></td>
        </tr>
      </table>
      <!--<p align="right"><img src="images/icon_excel.jpg" border="0" /> <a href="export_auction.asp">Export ALL</a></p>-->
      <p><font color="green"><%= strMessageText %></font></p>
      <table cellspacing="0" cellpadding="4" class="database_records">
        <tr class="innerdoctitle" align="center">
          <td>&nbsp;</td>
          <td>Lot</td>
          <td>Category</td>
          <td>Product</td>
          <td>Component(s)</td>
          <td>Comments</td>
          <td>LIC</td>
          <td>Reserve</td>
          <td>Winning bid</td>
          <td>Winner</td>
          <td>Sales order no</td>
          <td>Invoice no</td>          
          <td>&nbsp;</td>
        </tr>
        <%= strDisplayList %>
      </table></td>
  </tr>
</table>
</body>
</html>