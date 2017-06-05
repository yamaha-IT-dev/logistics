<% response.cookies("current_URL_cookie_logistics") = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?" & Request.Querystring %>
<!--#include file="include/connection_it.asp " -->
<!--#include file="class/clsComment.asp " -->
<% strSection = "transfer" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Update Transfer</title>
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="include/generic_form_validations.js"></script>
<script type="text/javascript" src="include/javascript.js"></script>
<script type="text/javascript" src="include/usableforms.js"></script>
<script language="JavaScript" type="text/javascript">
function popitup(url) {
	newwindow=window.open(url,'name','height=200,width=600');
	if (window.focus) {newwindow.focus()}
	return false;
}

function validateFormOnSubmit(theForm) {
	var reason = "";
	var blnSubmit = true;

	reason += validateEmptyField(theForm.txtInvoiceNo,"Invoice No");
	reason += validateSpecialCharacters(theForm.txtInvoiceNo,"Invoice No");

	if (theForm.cboReceived.value != 0) {
		reason += validateEmptyField(theForm.txtDateReceived,"Date Received");
	}
	
	if (theForm.cboBase.value == 0 && theForm.cboStatus.value == 0) {
		alert("BASE must be updated first");
		blnSubmit = false;
		return false;
	}
	

  	if (reason != "") {
    	alert("Oops... some fields need correction:\n" + reason);

		blnSubmit = false;
		return false;
  	}

	if (blnSubmit == true){
        theForm.Action.value = 'Update';
  		theForm.submit();

		return true;
    }
}

function submitComment(theForm) {
	var reason = "";
	var blnSubmit = true;
	
	reason += validateEmptyField(theForm.txtComment,"Comment");
	
	if (reason != "") {
    	alert("Some fields need correction:\n" + reason);

		blnSubmit = false;
		return false;
  	}
	
	if (blnSubmit == true){
		theForm.Action.value = 'Comment';
		
		return true;		
    }
}
</script>
<%

Sub getTransfer
	dim intQty1
	dim intPallet1
	dim intQty2
	dim intPallet2
	dim intQty3
	dim intPallet3
	dim intQty4
	dim intPallet4
	dim intQty5
	dim intPallet5
	dim intQty6
	dim intPallet6
	dim intQty7
	dim intPallet7
	dim intQty8
	dim intPallet8
	dim intQty9
	dim intPallet9
	dim intQty10
	dim intPallet10
	dim intQty11
	dim intPallet11
	dim intQty12
	dim intPallet12
	dim intQty13
	dim intPallet13
	dim intQty14
	dim intPallet14
	dim intQty15
	dim intPallet15
	dim intQty16
	dim intPallet16
	dim intQty17
	dim intPallet17
	dim intQty18
	dim intPallet18
	dim intQty19
	dim intPallet19
	dim intQty20
	dim intPallet20
	dim intTotalQty
	dim intTotalPallets

	dim intReceived1
	dim intReceived2
	dim intReceived3
	dim intReceived4
	dim intReceived5
	dim intReceived6
	dim intReceived7
	dim intReceived8
	dim intReceived9
	dim intReceived10
	dim intReceived11
	dim intReceived12
	dim intReceived13
	dim intReceived14
	dim intReceived15
	dim intReceived16
	dim intReceived17
	dim intReceived18
	dim intReceived19
	dim intReceived20

	dim intTotalReceived

	dim intID
	intID = request("id")

    call OpenDataBase()

	set rs = Server.CreateObject("ADODB.recordset")

	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic

	strSQL = "SELECT * FROM yma_transfer WHERE id = " & intID

	rs.Open strSQL, conn

	intQty1		= rs("qty_1")
	intPallet1 	= rs("pallet_1")
	intQty2 	= rs("qty_2")
	intPallet2 	= rs("pallet_2")
	intQty3 	= rs("qty_3")
	intPallet3 	= rs("pallet_3")
	intQty4 	= rs("qty_4")
	intPallet4 	= rs("pallet_4")
	intQty5 	= rs("qty_5")
	intPallet5 	= rs("pallet_5")
	intQty6 	= rs("qty_6")
	intPallet6 	= rs("pallet_6")
	intQty7 	= rs("qty_7")
	intPallet7 	= rs("pallet_7")
	intQty8 	= rs("qty_8")
	intPallet8 	= rs("pallet_8")
	intQty9 	= rs("qty_9")
	intPallet9 	= rs("pallet_9")
	intQty10 	= rs("qty_10")
	intPallet10 = rs("pallet_10")
	intQty11	= rs("qty_11")
	intPallet11 = rs("pallet_11")
	intQty12 	= rs("qty_12")
	intPallet12 = rs("pallet_12")
	intQty13 	= rs("qty_13")
	intPallet13 = rs("pallet_13")
	intQty14 	= rs("qty_14")
	intPallet14 = rs("pallet_14")
	intQty15 	= rs("qty_15")
	intPallet15 = rs("pallet_15")
	intQty16 	= rs("qty_16")
	intPallet16 = rs("pallet_16")
	intQty17 	= rs("qty_17")
	intPallet17 = rs("pallet_17")
	intQty18 	= rs("qty_18")
	intPallet18 = rs("pallet_18")
	intQty19 	= rs("qty_19")
	intPallet19 = rs("pallet_19")
	intQty20 	= rs("qty_20")
	intPallet20 = rs("pallet_20")

	intTotalQty = (intQty1 + intQty2 + intQty3 + intQty4 + intQty5 + intQty6 + intQty7 + intQty8 + intQty9 + intQty10 + intQty11 + intQty12 + intQty13 + intQty14 + intQty15 + intQty16 + intQty17 + intQty18 + intQty19 + intQty20)
	intTotalPallet = (intPallet1 + intPallet2 + intPallet3 + intPallet4 + intPallet5 + intPallet6 + intPallet7 + intPallet8 + intPallet9 + intPallet10 + intPallet11 + intPallet12 + intPallet13 + intPallet14 + intPallet15 + intPallet16 + intPallet17 + intPallet18 + intPallet19 + intPallet20)

	session("total_qty") 	= intTotalQty
	session("total_pallet") = intTotalPallet

	intReceived1 	= rs("received_1")
	intReceived2 	= rs("received_2")
	intReceived3 	= rs("received_3")
	intReceived4 	= rs("received_4")
	intReceived5 	= rs("received_5")
	intReceived6 	= rs("received_6")
	intReceived7 	= rs("received_7")
	intReceived8 	= rs("received_8")
	intReceived9 	= rs("received_9")
	intReceived10 	= rs("received_10")
	intReceived11 	= rs("received_11")
	intReceived12 	= rs("received_12")
	intReceived13 	= rs("received_13")
	intReceived14 	= rs("received_14")
	intReceived15 	= rs("received_15")
	intReceived16 	= rs("received_16")
	intReceived17 	= rs("received_17")
	intReceived18 	= rs("received_18")
	intReceived19 	= rs("received_19")
	intReceived20 	= rs("received_20")

	intTotalReceived = (intReceived1 + intReceived2 + intReceived3 + intReceived4 + intReceived5 + intReceived6 + intReceived7 + intReceived8 + intReceived9 + intReceived10 + intReceived11 + intReceived12 + intReceived13 + intReceived14 + intReceived15 + intReceived16 + intReceived17 + intReceived18 + intReceived19 + intReceived20)

	session("total_received") = intTotalReceived
	'Response.Write strSQL

    if not DB_RecSetIsEmpty(rs) Then
		session("priority") 	= rs("priority")
		session("warehouse") 	= rs("warehouse")
		session("product_1") 	= rs("product_1")
		session("qty_1") 		= rs("qty_1")
		session("pallet_1") 	= rs("pallet_1")
		session("info_1") 		= rs("info_1")
		session("received_1") 	= rs("received_1")
		session("product_2") 	= rs("product_2")
		session("qty_2") 		= rs("qty_2")
		session("pallet_2") 	= rs("pallet_2")
		session("info_2") 		= rs("info_2")
		session("received_2") 	= rs("received_2")
		session("product_3") 	= rs("product_3")
		session("qty_3") 		= rs("qty_3")
		session("pallet_3") 	= rs("pallet_3")
		session("info_3") 		= rs("info_3")
		session("received_3") 	= rs("received_3")
		session("product_4") 	= rs("product_4")
		session("qty_4") 		= rs("qty_4")
		session("pallet_4") 	= rs("pallet_4")
		session("info_4") 		= rs("info_4")
		session("received_4") 	= rs("received_4")
		session("product_5") 	= rs("product_5")
		session("qty_5") 		= rs("qty_5")
		session("pallet_5") 	= rs("pallet_5")
		session("info_5") 		= rs("info_5")
		session("received_5") 	= rs("received_5")
		session("product_6") 	= rs("product_6")
		session("qty_6") 		= rs("qty_6")
		session("pallet_6") 	= rs("pallet_6")
		session("info_6") 		= rs("info_6")
		session("received_6") 	= rs("received_6")
		session("product_7") 	= rs("product_7")
		session("qty_7") 		= rs("qty_7")
		session("pallet_7") 	= rs("pallet_7")
		session("info_7") 		= rs("info_7")
		session("received_7") 	= rs("received_7")
		session("product_8") 	= rs("product_8")
		session("qty_8") 		= rs("qty_8")
		session("pallet_8") 	= rs("pallet_8")
		session("info_8") 		= rs("info_8")
		session("received_8") 	= rs("received_8")
		session("product_9") 	= rs("product_9")
		session("qty_9") 		= rs("qty_9")
		session("pallet_9") 	= rs("pallet_9")
		session("info_9") 		= rs("info_9")
		session("received_9") 	= rs("received_9")
		session("product_10") 	= rs("product_10")
		session("qty_10") 		= rs("qty_10")
		session("pallet_10") 	= rs("pallet_10")
		session("info_10") 		= rs("info_10")
		session("received_10") 	= rs("received_10")
		session("product_11") 	= rs("product_11")
		session("qty_11") 		= rs("qty_11")
		session("pallet_11") 	= rs("pallet_11")
		session("info_11") 		= rs("info_11")
		session("received_11") 	= rs("received_11")
		session("product_12") 	= rs("product_12")
		session("qty_12") 		= rs("qty_12")
		session("pallet_12") 	= rs("pallet_12")
		session("info_12") 		= rs("info_12")
		session("received_12") 	= rs("received_12")
		session("product_13") 	= rs("product_13")
		session("qty_13") 		= rs("qty_13")
		session("pallet_13") 	= rs("pallet_13")
		session("info_13") 		= rs("info_13")
		session("received_13") 	= rs("received_13")
		session("product_14") 	= rs("product_14")
		session("qty_14") 		= rs("qty_14")
		session("pallet_14") 	= rs("pallet_14")
		session("info_14") 		= rs("info_14")
		session("received_14") 	= rs("received_14")
		session("product_15") 	= rs("product_15")
		session("qty_15") 		= rs("qty_15")
		session("pallet_15") 	= rs("pallet_15")
		session("info_15") 		= rs("info_15")
		session("received_15") 	= rs("received_15")
		session("product_16") 	= rs("product_16")
		session("qty_16") 		= rs("qty_16")
		session("pallet_16") 	= rs("pallet_16")
		session("info_16") 		= rs("info_16")
		session("received_16") 	= rs("received_16")
		session("product_17") 	= rs("product_17")
		session("qty_17") 		= rs("qty_17")
		session("pallet_17") 	= rs("pallet_17")
		session("info_17") 		= rs("info_17")
		session("received_17") 	= rs("received_17")
		session("product_18") 	= rs("product_18")
		session("qty_18") 		= rs("qty_18")
		session("pallet_18") 	= rs("pallet_18")
		session("info_18") 		= rs("info_18")
		session("received_18") 	= rs("received_18")
		session("product_19") 	= rs("product_19")
		session("qty_19") 		= rs("qty_19")
		session("pallet_19") 	= rs("pallet_19")
		session("info_19") 		= rs("info_19")
		session("received_19") 	= rs("received_19")
		session("product_20") 	= rs("product_20")
		session("qty_20") 		= rs("qty_20")
		session("pallet_20") 	= rs("pallet_20")
		session("info_20") 		= rs("info_20")
		session("received_20") 	= rs("received_20")
		session("transfer_date") = rs("transfer_date")
		session("date_received") = rs("date_received")
		session("transfer_comments") = rs("transfer_comments")
		session("status") 		= rs("status")
		session("pickup") 		= rs("pickup")
		session("pickup_date") 	= rs("pickup_date")
		session("pickup_time") 	= rs("pickup_time")
		session("booking_time") = rs("booking_time")
		session("received") 	= rs("received")
		session("invoice_no") 	= rs("invoice_no")
		session("date_created") = rs("date_created")
		session("created_by") 	= rs("created_by")
		session("date_modified")= rs("date_modified")
		session("modified_by") 	= rs("modified_by")
		session("comments") 	= rs("comments")
		session("base") 		= rs("base")
    end if

    call CloseDataBase()

end sub

sub updateTransfer

	dim strSQL
	dim intID
	intID = request("id")

	Call OpenDataBase()

	strSQL = "UPDATE yma_transfer SET "
	strSQL = strSQL & "received_1 = '" & Replace(Request.Form("txtReceived1"),"'","''") & "',"
	strSQL = strSQL & "received_2 = '" & Replace(Request.Form("txtReceived2"),"'","''") & "',"
	strSQL = strSQL & "received_3 = '" & Replace(Request.Form("txtReceived3"),"'","''") & "',"
	strSQL = strSQL & "received_4 = '" & Replace(Request.Form("txtReceived4"),"'","''") & "',"
	strSQL = strSQL & "received_5 = '" & Replace(Request.Form("txtReceived5"),"'","''") & "',"
	strSQL = strSQL & "received_6 = '" & Replace(Request.Form("txtReceived6"),"'","''") & "',"
	strSQL = strSQL & "received_7 = '" & Replace(Request.Form("txtReceived7"),"'","''") & "',"
	strSQL = strSQL & "received_8 = '" & Replace(Request.Form("txtReceived8"),"'","''") & "',"
	strSQL = strSQL & "received_9 = '" & Replace(Request.Form("txtReceived9"),"'","''") & "',"
	strSQL = strSQL & "received_10 = '" & Replace(Request.Form("txtReceived10"),"'","''") & "',"
	strSQL = strSQL & "received_11 = '" & Replace(Request.Form("txtReceived11"),"'","''") & "',"
	strSQL = strSQL & "received_12 = '" & Replace(Request.Form("txtReceived12"),"'","''") & "',"
	strSQL = strSQL & "received_13 = '" & Replace(Request.Form("txtReceived13"),"'","''") & "',"
	strSQL = strSQL & "received_14 = '" & Replace(Request.Form("txtReceived14"),"'","''") & "',"
	strSQL = strSQL & "received_15 = '" & Replace(Request.Form("txtReceived15"),"'","''") & "',"
	strSQL = strSQL & "received_16 = '" & Replace(Request.Form("txtReceived16"),"'","''") & "',"
	strSQL = strSQL & "received_17 = '" & Replace(Request.Form("txtReceived17"),"'","''") & "',"
	strSQL = strSQL & "received_18 = '" & Replace(Request.Form("txtReceived18"),"'","''") & "',"
	strSQL = strSQL & "received_19 = '" & Replace(Request.Form("txtReceived19"),"'","''") & "',"
	strSQL = strSQL & "received_20 = '" & Replace(Request.Form("txtReceived20"),"'","''") & "',"
	strSQL = strSQL & "invoice_no = '" & trim(Request.Form("txtInvoiceNo")) & "',"
	strSQL = strSQL & "comments = '" & Replace(Request.Form("txtComments"),"'","''") & "',"
	strSQL = strSQL & "date_modified = getdate(),"
	strSQL = strSQL & "modified_by = '" & session("UsrUserName") & "',"
	strSQL = strSQL & "pickup = '" & trim(Request.Form("cboPickup")) & "',"
	strSQL = strSQL & "booking_time = '" & trim(Request.Form("txtBookingTime")) & "',"
	strSQL = strSQL & "received = '" & trim(Request.Form("cboReceived")) & "',"
	strSQL = strSQL & "date_received = CONVERT(datetime,'" & trim(Request.Form("txtDateReceived")) & "',103),"
	strSQL = strSQL & "base = '" & trim(Request.Form("cboBase")) & "',"
	strSQL = strSQL & "status = '" & trim(Request.Form("cboStatus")) & "' WHERE id = " & intID

	'response.Write strSQL
	on error resume next
	conn.Execute strSQL

	if err <> 0 then
		strMessageText = err.description
	else
		strMessageText = "The transfer request has been updated."
	end if

	Call CloseDataBase()
end sub

sub main
	call UTL_validateLogin
	
	dim intID
	intID 	= request("id")
	
	call getTransfer
	call listComments(intID,transferModuleID)
	
	if Request.ServerVariables("REQUEST_METHOD") = "POST" then
		select case Trim(Request("Action"))
			case "Update"
				call updateTransfer
				call getTransfer
			case "Comment"
				call addComment(intID,transferModuleID)
				call listComments(intID,transferModuleID)	
		end select
	end if
	
end sub

call main

dim strMessageText
dim strCommentsList
%>
</head>
<body>
<table cellspacing="0" cellpadding="0" align="center" class="full_size_table" border="0">
  <!-- #include file="include/header.asp" -->
  <tr>
    <td class="first_content"><table cellpadding="5" cellspacing="0" border="0">
        <tr>
          <td><a href="list_transfer.asp"><img src="images/icon_transfer.jpg" border="0" alt="Transfer Requests" /></a></td>
          <td valign="top"><img src="images/backward_arrow.gif" border="0" /> <a href="list_transfer.asp">Back to List</a>
            <h2>Update Transfer</h2>
            <font color="green"><%= strMessageText %></font></td>
          <td valign="top"><table cellpadding="4" cellspacing="0" class="created_table">
              <tr>
                <td class="created_column_1"><strong>Created:</strong></td>
                <td class="created_column_2"><%= session("created_by") %></td>
                <td class="created_column_3"><%= displayDateFormatted(session("date_created")) %></td>
              </tr>
              <tr>
                <td><strong>Last modified:</strong></td>
                <td><%= session("modified_by") %></td>
                <td><%= displayDateFormatted(session("date_modified")) %></td>
              </tr>
            </table></td>
        </tr>
      </table>
      <form action="" method="post" name="form_update_transfer" id="form_update_transfer" onsubmit="return validateFormOnSubmit(this)">
        <table width="1024" border="0">
          <tr>
            <td valign="top" width="50%"><table cellpadding="5" cellspacing="0" class="item_maintenance_box">
                <tr>
                  <td class="item_maintenance_header">From - to : <%= session("warehouse") %> <% if len(session("priority")) = 1 then %>
                  <img src="images/icon_priority.gif" border="0" />
                  <% end if %></td>
                </tr>
                <tr>
                  <td valign="top"><table width="100%" cellpadding="3" cellspacing="0">
                      <tr>
                        <td>&nbsp;</td>
                        <td>Product</td>
                        <td>Qty</td>
                        <td>Pallet(s)</td>
                        <td>Info</td>
                        <td>Qty received</td>
                      </tr>
                      <tr class="highlighted_row">
                        <td>1:</td>
                        <td><%= session("product_1") %></td>
                        <td><%= session("qty_1") %></td>
                        <td><%= session("pallet_1") %></td>
                        <td><%= session("info_1") %></td>
                        <td><input type="text" id="txtReceived1" name="txtReceived1" maxlength="4" size="5" value="<%= session("received_1") %>" /></td>
                      </tr>
                      <% if session("product_2") <> "" then %>
                      <tr>
                        <td>2:</td>
                        <td><%= session("product_2") %></td>
                        <td><%= session("qty_2") %></td>
                        <td><%= session("pallet_2") %></td>
                        <td><%= session("info_2") %></td>
                        <td><input type="text" id="txtReceived2" name="txtReceived2" maxlength="4" size="5" value="<%= session("received_2") %>" /></td>
                      </tr>
                      <% end if
							   if session("product_3") <> "" then %>
                      <tr class="highlighted_row">
                        <td>3:</td>
                        <td><%= session("product_3") %></td>
                        <td><%= session("qty_3") %></td>
                        <td><%= session("pallet_3") %></td>
                        <td><%= session("info_3") %></td>
                        <td><input type="text" id="txtReceived3" name="txtReceived3" maxlength="4" size="5" value="<%= session("received_3") %>" /></td>
                      </tr>
                      <% end if
							   if session("product_4") <> "" then %>
                      <tr>
                        <td>4:</td>
                        <td><%= session("product_4") %></td>
                        <td><%= session("qty_4") %></td>
                        <td><%= session("pallet_4") %></td>
                        <td><%= session("info_4") %></td>
                        <td><input type="text" id="txtReceived4" name="txtReceived4" maxlength="4" size="5" value="<%= session("received_4") %>" /></td>
                      </tr>
                      <% end if
							   if session("product_5") <> "" then %>
                      <tr class="highlighted_row">
                        <td>5:</td>
                        <td><%= session("product_5") %></td>
                        <td><%= session("qty_5") %></td>
                        <td><%= session("pallet_5") %></td>
                        <td><%= session("info_5") %></td>
                        <td><input type="text" id="txtReceived5" name="txtReceived5" maxlength="4" size="5" value="<%= session("received_5") %>" /></td>
                      </tr>
                      <% end if
							   if session("product_6") <> "" then %>
                      <tr>
                        <td>6:</td>
                        <td><%= session("product_6") %></td>
                        <td><%= session("qty_6") %></td>
                        <td><%= session("pallet_6") %></td>
                        <td><%= session("info_6") %></td>
                        <td><input type="text" id="txtReceived6" name="txtReceived6" maxlength="4" size="5" value="<%= session("received_6") %>" /></td>
                      </tr>
                      <% end if
							   if session("product_7") <> "" then %>
                      <tr class="highlighted_row">
                        <td>7:</td>
                        <td><%= session("product_7") %></td>
                        <td><%= session("qty_7") %></td>
                        <td><%= session("pallet_7") %></td>
                        <td><%= session("info_7") %></td>
                        <td><input type="text" id="txtReceived7" name="txtReceived7" maxlength="4" size="5" value="<%= session("received_7") %>" /></td>
                      </tr>
                      <% end if
							   if session("product_8") <> "" then %>
                      <tr>
                        <td>8:</td>
                        <td><%= session("product_8") %></td>
                        <td><%= session("qty_8") %></td>
                        <td><%= session("pallet_8") %></td>
                        <td><%= session("info_8") %></td>
                        <td><input type="text" id="txtReceived8" name="txtReceived8" maxlength="4" size="5" value="<%= session("received_8") %>" /></td>
                      </tr>
                      <% end if
							   if session("product_9") <> "" then %>
                      <tr class="highlighted_row">
                        <td>9:</td>
                        <td><%= session("product_9") %></td>
                        <td><%= session("qty_9") %></td>
                        <td><%= session("pallet_9") %></td>
                        <td><%= session("info_9") %></td>
                        <td><input type="text" id="txtReceived9" name="txtReceived9" maxlength="4" size="5" value="<%= session("received_9") %>" /></td>
                      </tr>
                      <% end if
							   if session("product_10") <> "" then %>
                      <tr>
                        <td>10:</td>
                        <td><%= session("product_10") %></td>
                        <td><%= session("qty_10") %></td>
                        <td><%= session("pallet_10") %></td>
                        <td><%= session("info_10") %></td>
                        <td><input type="text" id="txtReceived10" name="txtReceived10" maxlength="4" size="5" value="<%= session("received_10") %>" /></td>
                      </tr>
                      <% end if %>
                      <% if session("product_11") <> "" then %>
                      <tr class="highlighted_row">
                        <td>11:</td>
                        <td><%= session("product_11") %></td>
                        <td><%= session("qty_11") %></td>
                        <td><%= session("pallet_11") %></td>
                        <td><%= session("info_11") %></td>
                        <td><input type="text" id="txtReceived11" name="txtReceived11" maxlength="4" size="5" value="<%= session("received_11") %>" /></td>
                      </tr>
                      <% end if
                               if session("product_12") <> "" then %>
                      <tr>
                        <td>12:</td>
                        <td><%= session("product_12") %></td>
                        <td><%= session("qty_12") %></td>
                        <td><%= session("pallet_12") %></td>
                        <td><%= session("info_12") %></td>
                        <td><input type="text" id="txtReceived12" name="txtReceived12" maxlength="4" size="5" value="<%= session("received_12") %>" /></td>
                      </tr>
                      <% end if
							   if session("product_13") <> "" then %>
                      <tr class="highlighted_row">
                        <td>13:</td>
                        <td><%= session("product_13") %></td>
                        <td><%= session("qty_13") %></td>
                        <td><%= session("pallet_13") %></td>
                        <td><%= session("info_13") %></td>
                        <td><input type="text" id="txtReceived13" name="txtReceived13" maxlength="4" size="5" value="<%= session("received_13") %>" /></td>
                      </tr>
                      <% end if
							   if session("product_14") <> "" then %>
                      <tr>
                        <td>14:</td>
                        <td><%= session("product_14") %></td>
                        <td><%= session("qty_14") %></td>
                        <td><%= session("pallet_14") %></td>
                        <td><%= session("info_14") %></td>
                        <td><input type="text" id="txtReceived14" name="txtReceived14" maxlength="4" size="5" value="<%= session("received_14") %>" /></td>
                      </tr>
                      <% end if
							   if session("product_15") <> "" then %>
                      <tr class="highlighted_row">
                        <td>15:</td>
                        <td><%= session("product_15") %></td>
                        <td><%= session("qty_15") %></td>
                        <td><%= session("pallet_15") %></td>
                        <td><%= session("info_15") %></td>
                        <td><input type="text" id="txtReceived15" name="txtReceived15" maxlength="4" size="5" value="<%= session("received_15") %>" /></td>
                      </tr>
                      <% end if
							   if session("product_16") <> "" then %>
                      <tr>
                        <td>16:</td>
                        <td><%= session("product_16") %></td>
                        <td><%= session("qty_16") %></td>
                        <td><%= session("pallet_16") %></td>
                        <td><%= session("info_16") %></td>
                        <td><input type="text" id="txtReceived16" name="txtReceived16" maxlength="4" size="5" value="<%= session("received_16") %>" /></td>
                      </tr>
                      <% end if
							   if session("product_17") <> "" then %>
                      <tr class="highlighted_row">
                        <td>17:</td>
                        <td><%= session("product_17") %></td>
                        <td><%= session("qty_17") %></td>
                        <td><%= session("pallet_17") %></td>
                        <td><%= session("info_17") %></td>
                        <td><input type="text" id="txtReceived17" name="txtReceived17" maxlength="4" size="5" value="<%= session("received_17") %>" /></td>
                      </tr>
                      <% end if
							   if session("product_18") <> "" then %>
                      <tr>
                        <td>18:</td>
                        <td><%= session("product_18") %></td>
                        <td><%= session("qty_18") %></td>
                        <td><%= session("pallet_18") %></td>
                        <td><%= session("info_18") %></td>
                        <td><input type="text" id="txtReceived18" name="txtReceived18" maxlength="4" size="5" value="<%= session("received_18") %>" /></td>
                      </tr>
                      <% end if
							   if session("product_19") <> "" then %>
                      <tr class="highlighted_row">
                        <td>19:</td>
                        <td><%= session("product_19") %></td>
                        <td><%= session("qty_19") %></td>
                        <td><%= session("pallet_19") %></td>
                        <td><%= session("info_19") %></td>
                        <td><input type="text" id="txtReceived19" name="txtReceived19" maxlength="4" size="5" value="<%= session("received_19") %>" /></td>
                      </tr>
                      <% end if
							   if session("product_20") <> "" then %>
                      <tr>
                        <td>20:</td>
                        <td><%= session("product_20") %></td>
                        <td><%= session("qty_20") %></td>
                        <td><%= session("pallet_20") %></td>
                        <td><%= session("info_20") %></td>
                        <td><input type="text" id="txtReceived20" name="txtReceived20" maxlength="4" size="5" value="<%= session("received_20") %>" /></td>
                      </tr>
                      <% end if %>
                      <tr>
                        <td>&nbsp;</td>
                        <td>Total:</td>
                        <td><u><%= session("total_qty") %></u></td>
                        <td><u><%= session("total_pallet") %></u></td>
                        <td>&nbsp;</td>
                        <td><u><%= session("total_received") %></u></td>
                      </tr>
                    </table>
                    <p align="right"><img src="images/icon_excel.jpg" border="0" /> <a href="export_transfer-product.asp?id=<%= request("id") %>">Export this product list</a></p></td>
                </tr>
                <tr>
                  <td>Delivery: <%= WeekDayName(WeekDay(session("pickup_date"))) %>, <%= FormatDateTime(session("pickup_date"),1) %>, <%= session("pickup_time") %></td>
                </tr>
                <tr>
                  <td valign="top"><img src="images/icon_quote.gif" border="0" /> <em><%= session("transfer_comments") %></em> <img src="images/icon_quote_end.gif" border="0" /></td>
                </tr>
              </table></td>
            <td valign="top" width="50%"><table cellpadding="5" cellspacing="0" class="item_maintenance_box">
                <tr>
                  <td colspan="2" class="item_maintenance_header">Paperwork</td>
                </tr>
                <tr>
                  <td width="25%">Invoice no<span class="mandatory">*</span>:</td>
                  <td width="75%"><input type="text" id="txtInvoiceNo" name="txtInvoiceNo" maxlength="20" size="30" value="<%= session("invoice_no") %>" /></td>
                </tr>
                <tr>
                  <td>Booking time:</td>
                  <td><input type="text" id="txtBookingTime" name="txtBookingTime" maxlength="6" size="8" value="<%= session("booking_time") %>" /></td>
                </tr>
                <tr>
                  <td>Picked up?</td>
                  <td><select name="cboPickup">
                      <option <% if session("pickup") = "0" then Response.Write " selected" end if%> value="0">No</option>
                      <option <% if session("pickup") = "1" then Response.Write " selected" end if%> value="1">Yes</option>
                    </select>
                    <% if session("pickup") = "1" and (session("warehouse") = "3K - TT" or session("warehouse") = "Excel - TT") or session("warehouse") = "YMA - TT" or session("warehouse") = "3L - TT" then %>
                    <img src="images/forward_arrow.gif" border="0" /> <a href="transfer_pickup-email-TT.asp">Notify Sam (TT Logistics)</a>
                    <% end if %>
                    <% if session("pickup") = "1" and (session("warehouse") = "TT - 3K" or session("warehouse") = "Excel - 3K") then %>
                    <img src="images/forward_arrow.gif" border="0" /> <a href="transfer_pickup-email-3K.asp">Notify Nicole (3K)</a>
                    <% end if %>
                    <% if session("pickup") = "1" and (session("warehouse") = "TT - Excel" or session("warehouse") = "3K - Excel" or session("warehouse") = "Excel - 3H") then %>
                    <img src="images/forward_arrow.gif" border="0" /> <a href="transfer_pickup-email-excel.asp">Notify Excel</a>
                    <% end if %></td>
                </tr>
                <tr>
                  <td>Received?</td>
                  <td><select name="cboReceived">
                      <option <% if session("received") = "0" then Response.Write " selected" end if%> value="0" rel="none">No</option>
                      <option <% if session("received") = "1" then Response.Write " selected" end if%> value="1" rel="received">Yes</option>
                    </select>
                    <% if session("received") = "1" then %>
                    <img src="images/forward_arrow.gif" border="0" /> <a href="transfer_received-email.asp">Notify Yamaha Logistics</a>
                    <% end if %></td>
                </tr>
                <tr rel="received">
                  <td>Date received<span class="mandatory">*</span>:</td>
                  <td><input type="text" id="txtDateReceived" name="txtDateReceived" maxlength="10" size="10" value="<%= session("date_received") %>" />
                    <em>DD/MM/YYYY</em></td>
                </tr>
                <tr>
                  <td>Base updated?</td>
                  <td><select name="cboBase">
                      <option <% if session("base") = "0" then Response.Write " selected" end if%> value="0">No</option>
                      <option <% if session("base") = "1" then Response.Write " selected" end if%> value="1">Yes</option>
                    </select>
                    <% if session("base") = "1" and (session("warehouse") = "3K - TT" or session("warehouse") = "Excel - TT") or session("warehouse") = "YMA - TT" or session("warehouse") = "3L - TT" then %>
                    <img src="images/forward_arrow.gif" border="0" /> <a href="transfer_base-updated-TT.asp">Notify Sam (TT Logistics)</a>
                    <% end if %>
                    <% if session("base") = "1" and (session("warehouse") = "TT - 3K" or session("warehouse") = "Excel - 3K") then %>
                    <img src="images/forward_arrow.gif" border="0" /> <a href="transfer_base-updated-3K.asp">Notify Nicole (3K)</a>
                    <% end if %>
                    <% if session("base") = "1" and (session("warehouse") = "TT - Excel" or session("warehouse") = "3K - Excel" or session("warehouse") = "Excel - 3H") then %>
                    <img src="images/forward_arrow.gif" border="0" /> <a href="transfer_base-updated-excel.asp">Notify Excel</a>
                    <% end if %></td>
                </tr>
                <tr>
                  <td>Comments:</td>
                  <td><textarea name="txtComments" id="txtComments" cols="40" rows="3"><%= session("comments") %></textarea></td>
                </tr>
                <tr class="status_row">
                  <td>Status:</td>
                  <td><select name="cboStatus">
                      <option <% if session("status") = "1" then Response.Write " selected" end if%> value="1">Open</option>
                      <option <% if session("status") = "0" then Response.Write " selected" end if%> value="0">Completed</option>
                    </select></td>
                </tr>                
              </table>
              <p>
                <input type="hidden" name="Action" />
                <input type="submit" value="Update Transfer" />
              </p></td>
          </tr>
        </table>
      </form>
      <h2>Comments<br />
        <img src="images/comment_bar.jpg" border="0" /></h2>
      <table cellpadding="5" cellspacing="0" border="0" class="comments_box">
        <%= strCommentsList %>
        <tr>
          <td><form action="" method="post" onsubmit="return submitComment(this)">
              <p>
                <input type="text" name="txtComment" id="txtComment" maxlength="60" size="65" />
                <input type="hidden" name="Action" />
                <input type="submit" value="Add Comment" />
              </p>
            </form></td>
        </tr>
      </table></td>
  </tr>
</table>
<script type="text/javascript" src="include/moment.js"></script>
<script type="text/javascript" src="include/pikaday.js"></script>
<script type="text/javascript">	
	var picker = new Pikaday(
    {
        field: document.getElementById('txtDateReceived'),		
        firstDay: 1,
        minDate: new Date('2010-01-01'),
        maxDate: new Date('2030-12-31'),
        yearRange: [2013,2030],
		format: 'DD/MM/YYYY'
    });			
</script>
</body>
</html>