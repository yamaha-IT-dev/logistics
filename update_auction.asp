<% response.cookies("current_URL_cookie_logistics") = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?" & Request.Querystring %>
<!--#include file="include/connection_it.asp " -->
<% strSection = "auction" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Update Auction</title>
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="include/generic_form_validations.js"></script>
<script language="JavaScript" type="text/javascript">
function validateFormOnSubmit(theForm) {
	var reason = "";
	var blnSubmit = true;

	reason += validateEmptyField(theForm.txtProduct,"Product");
	
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
<%

Sub getAuctionItem

	dim intID
	intID = request("id")

    call OpenDataBase()

	set rs = Server.CreateObject("ADODB.recordset")

	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic

	strSQL = "SELECT * FROM yma_auction_aug2012 WHERE id = " & intID

	rs.Open strSQL, conn

	'Response.Write strSQL

    if not DB_RecSetIsEmpty(rs) Then
		session("category") = rs("category")
		session("product") = rs("product")
		session("basecode") = rs("basecode")
		session("lic") = rs("lic")
		session("winner") = rs("winner")
		session("bids") = rs("bids")
		session("reserve") = rs("reserve")
		session("winning_bid") = rs("winning_bid")
		session("status") = rs("status")
		session("date_modified") = rs("date_modified")
		session("modified_by") = rs("modified_by")
		session("sales_order_no") = rs("sales_order_no")
		session("invoice_no") = rs("invoice_no")
		session("comments") = rs("comments")
    end if

    call CloseDataBase()

end sub

sub updateAuctionItem

	dim strSQL
	dim intID
	intID = request("id")

	Call OpenDataBase()

	strSQL = "UPDATE yma_auction_aug2012 SET "
	strSQL = strSQL & "product = '" & Replace(Request.Form("txtProduct"),"'","''") & "',"
	strSQL = strSQL & "sales_order_no = '" & Replace(Request.Form("txtSalesOrderNo"),"'","''") & "',"
	strSQL = strSQL & "invoice_no = '" & Replace(Request.Form("txtInvoiceNo"),"'","''") & "',"
	strSQL = strSQL & "comments = '" & Replace(Request.Form("txtComments"),"'","''") & "',"
	strSQL = strSQL & "date_modified = getdate(),"
	strSQL = strSQL & "modified_by = '" & session("UsrUserName") & "' WHERE id = " & intID

	'response.Write strSQL
	on error resume next
	conn.Execute strSQL

	On error Goto 0

	if err <> 0 then
		strMessageText = err.description
	else
		strMessageText = "The item has been updated."
	end if

	Call CloseDataBase()
end sub

sub backRecordButton
	dim strSQL
	dim intID
	intID = request("id")

	dim strDisplayBackButton

	set rs = Server.CreateObject("ADODB.recordset")

	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic

	Call OpenDataBase()
	strSQL = "SELECT TOP 1 id FROM yma_auction_aug2012 WHERE id < '" & intID & "' ORDER BY id DESC"

	rs.Open strSQL, conn

	if not DB_RecSetIsEmpty(rs) Then
		session("previous_button") = "<a href=""update_auction.asp?id=" & rs("id") & """><img src=""images/backpage.png"" border=""0"" /></a>"
    end if

	call CloseDataBase()
end sub

sub nextRecordButton
	dim strSQL
	dim intID
	intID = request("id")

	dim strDisplayNextButton

	set rs = Server.CreateObject("ADODB.recordset")

	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic

	Call OpenDataBase()
	strSQL = "SELECT TOP 1 id FROM yma_auction_aug2012 WHERE id > '" & intID & "' ORDER BY id"

	rs.Open strSQL, conn

	if not DB_RecSetIsEmpty(rs) Then
		session("next_button") = "<a href=""update_auction.asp?id=" & rs("id") & """><img src=""images/nextpage.png"" border=""0"" /></a>"
    end if

	call CloseDataBase()
end sub

sub main
	call UTL_validateLogin
	
	call backRecordButton
	call nextRecordButton
	
	call getAuctionItem
	
	if Request.ServerVariables("REQUEST_METHOD") = "POST" then
		if Trim(Request("Action")) = "Update" then
			call updateAuctionItem
			call getAuctionItem		
		end if
	end if
end sub

call main

dim strMessageText
%>
</head>
<body>
<form action="" method="post" name="form_update_action" id="form_update_action" onsubmit="return validateFormOnSubmit(this)">
  <table width="100%" cellpadding="0" cellspacing="0">
    <!-- #include file="include/header.asp" -->
    <tr>
      <td class="first_content">
      <table cellpadding="5" cellspacing="0" border="0">
        <tr>
          <td valign="top"><a href="list_auctions.asp"><img src="../auction/images/ybay_logo_clean.jpg" border="0" alt="yBay" /></a></td>
          <td valign="top"><img src="images/backward_arrow.gif" border="0" /> <a href="list_auctions.asp">Back to List</a>
            <h2>Update Auction Winner</h2>
            <font color="green"><%= strMessageText %></font></td>
        </tr>
      </table>
        <table cellpadding="4" cellspacing="0" class="created_table">
          <tr>
            <td class="created_column_1"><strong>Last modified:</strong></td>
            <td class="created_column_2"><%= session("modified_by") %></td>
            <td class="created_column_3"><%= displayDateFormatted(session("date_modified")) %></td>
          </tr>
        </table>
        <br />
        <table cellpadding="5" cellspacing="0" class="item_maintenance_box">
          <tr>
            <td colspan="2" class="item_maintenance_header">Lot no: <%= request("id") %></td>
          </tr>
          <tr>
            <td width="30%">Category:</td>
            <td width="70%"><%= session("category") %></td>
          </tr>
          <tr>
            <td>Product<span class="mandatory">*</span>:</td>
            <td><input type="text" id="txtProduct" name="txtProduct" maxlength="70" size="50" value="<%= session("product") %>" /></td>
          </tr>
          <tr>
            <td>Component(s):</td>
            <td><%= session("basecode") %></td>
          </tr>
          <tr>
            <td>LIC:</td>
            <td>$<%= session("lic") %></td>
          </tr>
          <tr>
            <td>Reserve:</td>
            <td>$<%= session("reserve") %></td>
          </tr>
          <tr>
            <td>Bids:</td>
            <td><%= session("bids") %></td>
          </tr>
          <tr class="status_row">
            <td>Winner:</td>
            <td><%= session("winner") %></td>
          </tr>
          <tr class="status_row">
            <td>Winning bid:</td>
            <td>$<%= session("winning_bid") %></td>
          </tr>
          <tr>
            <td>Sales order no:</td>
            <td><input type="text" id="txtSalesOrderNo" name="txtSalesOrderNo" maxlength="20" size="20" value="<%= session("sales_order_no") %>" /></td>
          </tr>
          <tr>
            <td>Invoice no:</td>
            <td><input type="text" id="txtInvoiceNo" name="txtInvoiceNo" maxlength="20" size="20" value="<%= session("invoice_no") %>" /></td>
          </tr>
          <tr>
            <td>Comments:</td>
            <td><input type="text" id="txtComments" name="txtComments" maxlength="90" size="50" value="<%= session("comments") %>" /></td>
          </tr>
          <tr>
            <td colspan="2"><input type="hidden" name="Action" />
              <input type="submit" value="Update" /></td>
          </tr>
        </table>
        <table width="500">
          <tr>
            <td align="left"><%= session("previous_button") %></td>
            <td align="right"><%= session("next_button") %></td>
          </tr>
        </table>
      </td>
    </tr>
    
  </table>
</form>
</body>
</html>