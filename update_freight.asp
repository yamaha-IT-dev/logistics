<% response.cookies("current_URL_cookie_logistics") = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?" & Request.Querystring %>
<!--#include file="include/connection_it.asp " -->
<!--#include file="class/clsComment.asp " -->
<% strSection = "freight" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Update Freight Request</title>
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="include/generic_form_validations.js"></script>
<script language="JavaScript" type="text/javascript">
function validateFormOnSubmit(theForm) {
	var reason = "";
	var blnSubmit = true;

	reason += validateEmptyField(theForm.txtConnote,"Con-note");
	reason += validateSpecialCharacters(theForm.txtConnote,"Con-note");

  	if (reason != "") {
    	alert("Some fields need correction:\n" + reason);

		blnSubmit = false;
		return false;
  	}

	if (blnSubmit == true){
        theForm.Action.value = 'Update';  		

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

Sub getFreight

	dim intID
	intID = request("id")

    call OpenDataBase()

	set rs = Server.CreateObject("ADODB.recordset")

	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic

	strSQL = "SELECT * FROM yma_freight WHERE id = " & intID

	rs.Open strSQL, conn

    if not DB_RecSetIsEmpty(rs) Then
		session("email") 			= Lcase(rs("email"))
		session("priority") 		= rs("priority")
		session("supplier") 		= rs("supplier")
		session("pickup") 			= rs("pickup")
		session("division") 		= rs("division")
		session("quote_required") 	= rs("quote_required")
		session("amount") 			= rs("amount")
		
		'PICKUP
		session("pickup_name") 		= rs("pickup_name")
		session("pickup_contact") 	= rs("pickup_contact")
		session("pickup_phone") 	= rs("pickup_phone")
		session("pickup_address") 	= rs("pickup_address")
		session("pickup_city") 		= rs("pickup_city")
		session("pickup_state") 	= rs("pickup_state")
		session("pickup_postcode") 	= rs("pickup_postcode")
		session("pickup_date") 		= rs("pickup_date")
		session("pickup_time") 		= rs("pickup_time")
		session("pickup_comments") 	= rs("pickup_comments")
		session("return_pickup") 	= rs("return_pickup")
		session("return_pickup_date") = rs("return_pickup_date")
		session("return_pickup_time") = rs("return_pickup_time")

		'RECEIVER
		session("receiver_name") 	= rs("receiver_name")
		session("receiver_contact") = rs("receiver_contact")
		session("receiver_phone") 	= rs("receiver_phone")
		session("receiver_address") = rs("receiver_address")
		session("receiver_city") 	= rs("receiver_city")
		session("receiver_state") 	= rs("receiver_state")
		session("receiver_postcode") = rs("receiver_postcode")
		session("delivery_date") 	= rs("delivery_date")
		session("delivery_time") 	= rs("delivery_time")
		session("receiver_comments") = rs("receiver_comments")

		'ITEM 1
		session("description") 	= rs("description")
		session("item_ref") 	= rs("item_ref")
		session("items") 		= rs("items")
		session("pallets") 		= rs("pallets")
		session("length") 		= rs("length")
		session("width") 		= rs("width")
		session("height") 		= rs("height")
		session("weight") 		= rs("weight")

		'ITEM 2
		session("description2") = rs("description2")
		session("item_ref2") 	= rs("item_ref2")
		session("items2") 		= rs("items2")
		session("pallets2") 	= rs("pallets2")
		session("length2") 		= rs("length2")
		session("width2") 		= rs("width2")
		session("height2") 		= rs("height2")
		session("weight2") 		= rs("weight2")

		'ITEM 3
		session("description3") = rs("description3")
		session("item_ref3") 	= rs("item_ref3")
		session("items3") 		= rs("items3")
		session("pallets3") 	= rs("pallets3")
		session("length3") 		= rs("length3")
		session("width3") 		= rs("width3")
		session("height3") 		= rs("height3")
		session("weight3") 		= rs("weight3")

		'ITEM 4
		session("description4") = rs("description4")
		session("item_ref4") 	= rs("item_ref4")
		session("items4") 		= rs("items4")
		session("pallets4") 	= rs("pallets4")
		session("length4") 		= rs("length4")
		session("width4") 		= rs("width4")
		session("height4") 		= rs("height4")
		session("weight4") 		= rs("weight4")

		'ITEM 5
		session("description5") = rs("description5")
		session("item_ref5") 	= rs("item_ref5")
		session("items5") 		= rs("items5")
		session("pallets5") 	= rs("pallets5")
		session("length5") 		= rs("length5")
		session("width5") 		= rs("width5")
		session("height5") 		= rs("height5")
		session("weight5") 		= rs("weight5")

		'ITEM 6
		session("description6") = rs("description6")
		session("item_ref6") 	= rs("item_ref6")
		session("items6") 		= rs("items6")
		session("pallets6") 	= rs("pallets6")
		session("length6") 		= rs("length6")
		session("width6") 		= rs("width6")
		session("height6") 		= rs("height6")
		session("weight6") 		= rs("weight6")

		'ITEM 7
		session("description7") = rs("description7")
		session("item_ref7") 	= rs("item_ref7")
		session("items7") 		= rs("items7")
		session("pallets7") 	= rs("pallets7")
		session("length7") 		= rs("length7")
		session("width7") 		= rs("width7")
		session("height7") 		= rs("height7")
		session("weight7") 		= rs("weight7")

		'ITEM 8
		session("description8") = rs("description8")
		session("item_ref8") 	= rs("item_ref8")
		session("items8") 		= rs("items8")
		session("pallets8") 	= rs("pallets8")
		session("length8") 		= rs("length8")
		session("width8") 		= rs("width8")
		session("height8") 		= rs("height8")
		session("weight8") 		= rs("weight8")

		session("comments") 		= rs("comments")
		session("consignment_no") 	= rs("consignment_no")
		session("return_connote") 	= rs("return_connote")
		session("status") 			= rs("status")
		session("date_created") 	= rs("date_created")
		session("created_by") 		= rs("created_by")
		session("date_modified") 	= rs("date_modified")
		session("modified_by") 		= rs("modified_by")
    end if

    call CloseDataBase()
end sub

sub updateFreight
	dim strSQL
	dim intID

	intID = request("id")

	Call OpenDataBase()

	strSQL = "UPDATE yma_freight SET "
	strSQL = strSQL & "supplier = '" & trim(Request.Form("cboSupplier")) & "',"
	strSQL = strSQL & "pickup = '" & trim(Request.Form("cboPickup")) & "',"
	strSQL = strSQL & "consignment_no = '" & trim(Request.Form("txtConnote")) & "',"
	strSQL = strSQL & "return_connote = '" & trim(Request.Form("txtReturnConnote")) & "',"
	strSQL = strSQL & "comments = '" & Replace(Request.Form("txtComments"),"'","''") & "',"
	strSQL = strSQL & "date_modified = getdate(),"
	strSQL = strSQL & "modified_by = '" & session("UsrUserName") & "',"
	strSQL = strSQL & "status = '" & trim(Request.Form("cboStatus")) & "' WHERE id = " & intID

	'response.Write strSQL
	on error resume next
	conn.Execute strSQL

	if err <> 0 then
		strMessageText = err.description
	else
		strMessageText = "The freight request has been updated."
	end if

	'if Request.Form("cboStatus") = 0 then
	' 	call sendEmail
	'end if

	Call CloseDataBase()
end sub

'-----------------------------------------------------------------------------
' SEND EMAIL notification to the requester when the freight has been completed
'-----------------------------------------------------------------------------

sub sendEmail

	Dim objCDOSYSMail
	Set objCDOSYSMail = Server.CreateObject("CDO.Message")
	Dim objCDOSYSCnfg
	Set objCDOSYSCnfg = Server.CreateObject("CDO.Configuration")

	Set oMail = Server.CreateObject("CDO.Message")
	Set iConf = Server.CreateObject("CDO.Configuration")
	Set Flds = iConf.Fields

	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 1
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "172.29.64.13"
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 10
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
	iConf.Fields.Update

	objCDOSYSMail.Configuration = objCDOSYSCnfg

	emailTo = trim(session("email"))
	emailCc = "logistics-aus@music.yamaha.com"
	emailBcc = "Matthew.Madden@music.yamaha.com"

	emailFrom 		= "automailer@music.yamaha.com"
	emailSubject 	= "Your Freight Request has been completed"

	emailBodyText   = 	"Hi there," & vbCrLf _
					&	"" & vbCrLf _
					&	"Your freight request created at " & session("date_created") & " to " & session("receiver_name")  & " has been completed. Thank you." & vbCrLf _
					&	"" & vbCrLf _
					&	"Regards," & vbCrLf _
					&	""  & vbCrLf _
					&	"Yamaha Logistics Division" & vbCrLf _
					&	""  & vbCrLf _
					&   "This is an automated email - please do not reply to this email."

	Set oMail.Configuration = iConf
	oMail.To 		= emailTo
	oMail.Cc		= emailCc
	oMail.Bcc		= emailBcc
	oMail.From 		= emailFrom
	oMail.Subject 	= emailSubject
	oMail.TextBody 	= emailBodyText
	oMail.Send

	Set iConf = Nothing
	Set Flds = Nothing

	strEmailMessageText = "An email has been sent to the requester."
end sub

sub main
	call UTL_validateLogin
	
	dim intID
	intID 	= request("id")		
	
	if Request.ServerVariables("REQUEST_METHOD") = "POST" then
		select case Trim(Request("Action"))
			case "Update"
				call updateFreight				
			case "Comment"
				call addComment(intID,freightModuleID)				
		end select
	end if
	
	call getFreight
	call listComments(intID,freightModuleID)
end sub

call main

dim strMessageText
dim strEmailMessageText
dim strCommentsList
%>
</head>
<body>
<table cellspacing="0" cellpadding="0" align="center" class="full_size_table" border="0">
  <!-- #include file="include/header.asp" -->
  <tr>
    <td class="first_content"><table cellpadding="5" cellspacing="0" border="0">
        <tr>
          <td valign="top"><a href="list_freights.asp"><img src="images/icon_freight.jpg" border="0" alt="Freight Requests" /></a></td>
          <td valign="top"><img src="images/backward_arrow.gif" width="6" height="12" border="0" /> <a href="list_freights.asp">Back to List</a>
            <h2>Update Freight Request</h2>
            <font color="green"><%= strMessageText %></font></td>
          <td valign="top"><table cellpadding="4" cellspacing="0" class="created_table_wide">
                <tr>
                  <td width="20%"><strong>Created:</strong></td>
                  <td width="40%"><%= session("email") %> <% if len(session("priority")) = 1 then %>
                  <img src="images/icon_priority.gif" border="0" />
                  <% end if %></td>
                  <td width="40%"><%= displayDateFormatted(session("date_created")) %></td>
                </tr>
                <tr>
                  <td>&nbsp;</td>
                  <td><%= session("division") %></td>
                  <td><strong>Quote?</strong> <%= session("quote_required") %> $<%= session("amount") %></td>
                </tr>
                <tr>
                  <td><strong>Last modified:</strong></td>
                  <td><%= session("modified_by") %></td>
                  <td><%= displayDateFormatted(session("date_modified")) %></td>
                </tr>
              </table></td>
        </tr>
      </table>
      <font color="red"><%= strEmailMessageText %></font>
      <form action="" method="post" name="form_update_freight" id="form_update_freight" onsubmit="return validateFormOnSubmit(this)">
        <table border="0" width="800">
          <tr>
            <td valign="top"><table width="450" cellpadding="3" cellspacing="0" class="thin_border">
                <tr>
                  <td class="item_maintenance_header">PICKUP</td>
                </tr>
                <tr>
                  <td align="center"><b><%= session("pickup_name") %></b></td>
                </tr>
                <tr>
                  <td align="center"><%= session("pickup_contact") %> - <%= session("pickup_phone") %></td>
                </tr>
                <tr>
                  <td valign="top" align="center"><%= session("pickup_address") %><br />
                    <%= session("pickup_city") %>&nbsp;<%= session("pickup_state") %>&nbsp;<%= session("pickup_postcode") %></td>
                </tr>
                <tr>
                  <td align="center"><b><%= WeekDayName(WeekDay(session("pickup_date"))) %>, <%= FormatDateTime(session("pickup_date"),1) %></b> at <%= session("pickup_time") %></td>
                </tr>
                <tr>
                  <td valign="top" align="center"><img src="images/icon_quote.gif" border="0" /> <em><%= session("pickup_comments") %></em></td>
                </tr>
              </table>
              <br />
              <% if session("return_pickup") = 1 then %>
              <table width="450" cellpadding="3" cellspacing="0" class="thin_border_red">
                <tr>
                  <td colspan="2" bgcolor="#FF0000" style="color:#FFF"><strong>Return to Pickup Details</strong></td>
                </tr>
                <tr>
                  <td width="30%">Date / Time:</td>
                  <td width="70%"><%= session("return_pickup_date") %> - <%= session("return_pickup_time") %></td>
                </tr>
                <tr>
                  <td>Return Connote:</td>
                  <td><input type="text" id="txtReturnConnote" name="txtReturnConnote" maxlength="20" size="30" value="<%= session("return_connote") %>" /></td>
                </tr>
              </table>
              <% end if %></td>
            <td valign="top"><table border="0" width="450" cellpadding="3" cellspacing="0" class="thin_border">
                <tr>
                  <td class="item_maintenance_header">RECEIVER</td>
                </tr>
                <tr>
                  <td align="center"><b><%= session("receiver_name") %></b></td>
                </tr>
                <tr>
                  <td align="center"><%= session("receiver_contact") %> - <%= session("receiver_phone") %></td>
                </tr>
                <tr>
                  <td valign="top" align="center"><%= session("receiver_address") %><br />
                    <%= session("receiver_city") %>&nbsp;<%= session("receiver_state") %>&nbsp;<%= session("receiver_postcode") %></td>
                </tr>
                <tr>
                  <td align="center"><b><%= WeekDayName(WeekDay(session("delivery_date"))) %>, <%= FormatDateTime(session("delivery_date"),1) %></b> at <%= session("delivery_time") %></td>
                </tr>
                <tr>
                  <td valign="top" align="center"><img src="images/icon_quote.gif" border="0" /> <em><%= session("receiver_comments") %></em></td>
                </tr>
              </table></td>
          </tr>
          <tr>
            <td valign="top"><table width="450" cellpadding="3" cellspacing="0" class="thin_border">
                <tr>
                  <td colspan="2" class="item_maintenance_header">Line 1</td>
                </tr>
                <tr>
                  <td width="30%">Description:</td>
                  <td width="70%"><%= session("description") %></td>
                </tr>
                <tr>
                  <td>Ref:</td>
                  <td><%= session("item_ref") %></td>
                </tr>
                <tr>
                  <td>Qty:</td>
                  <td><%= session("items") %></td>
                </tr>
                <tr>
                  <td>Pallets:</td>
                  <td><%= session("pallets") %></td>
                </tr>
                <tr>
                  <td>L / W / H:</td>
                  <td><%= session("length") %> cm / <%= session("width") %> cm / <%= session("height") %> cm</td>
                </tr>
                <tr>
                  <td>&nbsp;</td>
                  <td><%= session("weight") %> kg</td>
                </tr>
              </table>
              <br />
              <% if len(session("description3")) > 0 or len(session("item_ref3")) > 0 or len(session("items3")) > 0 or len(session("pallets3")) > 0 or len(session("length3")) > 0 or len(session("width3")) > 0 or len(session("height3")) > 0 or len(session("weight3")) > 0 then %>
              <table width="450" cellpadding="3" cellspacing="0" class="thin_border">
                <tr>
                  <td colspan="2" class="item_maintenance_header">Line 3</td>
                </tr>
                <tr>
                  <td width="30%">Description:</td>
                  <td width="70%"><%= session("description3") %></td>
                </tr>
                <tr>
                  <td>Ref:</td>
                  <td><%= session("item_ref3") %></td>
                </tr>
                <tr>
                  <td>Qty:</td>
                  <td><%= session("items3") %></td>
                </tr>
                <tr>
                  <td>Pallets:</td>
                  <td><%= session("pallets3") %></td>
                </tr>
                <tr>
                  <td>L / W / H:</td>
                  <td><%= session("length3") %> cm / <%= session("width3") %> cm / <%= session("height3") %> cm</td>
                </tr>
                <tr>
                  <td>&nbsp;</td>
                  <td><%= session("weight3") %> kg</td>
                </tr>
              </table>
              <% end if %>
              <% if len(session("description5")) > 0 or len(session("item_ref5")) > 0 or len(session("items5")) > 0 or len(session("pallets5")) > 0 or len(session("length5")) > 0 or len(session("width5")) > 0 or len(session("height5")) > 0 or len(session("weight5")) > 0 then %>
              <br />
              <table width="450" cellpadding="3" cellspacing="0" class="thin_border">
                <tr>
                  <td colspan="2" class="item_maintenance_header">Line 5</td>
                </tr>
                <tr>
                  <td width="30%">Description:</td>
                  <td width="70%"><%= session("description5") %></td>
                </tr>
                <tr>
                  <td>Ref:</td>
                  <td><%= session("item_ref5") %></td>
                </tr>
                <tr>
                  <td>Qty:</td>
                  <td><%= session("items5") %></td>
                </tr>
                <tr>
                  <td>Pallets:</td>
                  <td><%= session("pallets5") %></td>
                </tr>
                <tr>
                  <td>L / W / H:</td>
                  <td><%= session("length5") %> cm / <%= session("width5") %> cm / <%= session("height5") %> cm</td>
                </tr>
                <tr>
                  <td>&nbsp;</td>
                  <td><%= session("weight5") %> kg</td>
                </tr>
              </table>
              <% end if %>
              <% if len(session("description7")) > 0 or len(session("item_ref7")) > 0 or len(session("items7")) > 0 or len(session("pallets7")) > 0 or len(session("length7")) > 0 or len(session("width7")) > 0 or len(session("height7")) > 0 or len(session("weight7")) > 0 then %>
              <br />
              <table width="450" cellpadding="3" cellspacing="0" class="thin_border">
                <tr>
                  <td colspan="2" class="item_maintenance_header">Line 7</td>
                </tr>
                <tr>
                  <td width="30%">Description:</td>
                  <td width="70%"><%= session("description7") %></td>
                </tr>
                <tr>
                  <td>Ref:</td>
                  <td><%= session("item_ref7") %></td>
                </tr>
                <tr>
                  <td>Qty:</td>
                  <td><%= session("items7") %></td>
                </tr>
                <tr>
                  <td>Pallets:</td>
                  <td><%= session("pallets7") %></td>
                </tr>
                <tr>
                  <td>L / W / H:</td>
                  <td><%= session("length7") %> cm / <%= session("width7") %> cm / <%= session("height7") %> cm</td>
                </tr>
                <tr>
                  <td>&nbsp;</td>
                  <td><%= session("weight7") %> kg</td>
                </tr>
              </table>
              <% end if %></td>
            <td valign="top"><% if len(session("description2")) > 0 or len(session("item_ref2")) > 0 or len(session("items2")) > 0 or len(session("pallets2")) > 0 or len(session("length2")) > 0 or len(session("width2")) > 0 or len(session("height2")) > 0 or len(session("weight2")) > 0 then %>
              <table width="450" cellpadding="3" cellspacing="0" class="thin_border">
                <tr>
                  <td colspan="2" class="item_maintenance_header">Line 2</td>
                </tr>
                <tr>
                  <td width="30%">Description:</td>
                  <td width="70%"><%= session("description2") %></td>
                </tr>
                <tr>
                  <td>Ref:</td>
                  <td><%= session("item_ref2") %></td>
                </tr>
                <tr>
                  <td>Qty:</td>
                  <td><%= session("items2") %></td>
                </tr>
                <tr>
                  <td>Pallets:</td>
                  <td><%= session("pallets2") %></td>
                </tr>
                <tr>
                  <td>L / W / H:</td>
                  <td><%= session("length2") %> cm / <%= session("width2") %> cm / <%= session("height2") %> cm</td>
                </tr>
                <tr>
                  <td>&nbsp;</td>
                  <td><%= session("weight2") %> kg</td>
                </tr>
              </table>
              <% end if %>
              <% if len(session("description4")) > 0 or len(session("item_ref4")) > 0 or len(session("items4")) > 0 or len(session("pallets4")) > 0 or len(session("length4")) > 0 or len(session("width4")) > 0 or len(session("height4")) > 0 or len(session("weight4")) > 0 then %>
              <br />
              <table width="450" cellpadding="3" cellspacing="0" class="thin_border">
                <tr>
                  <td colspan="2" class="item_maintenance_header">Line 4</td>
                </tr>
                <tr>
                  <td width="30%">Description:</td>
                  <td width="70%"><%= session("description4") %></td>
                </tr>
                <tr>
                  <td>Ref:</td>
                  <td><%= session("item_ref4") %></td>
                </tr>
                <tr>
                  <td>Qty:</td>
                  <td><%= session("items4") %></td>
                </tr>
                <tr>
                  <td>Pallets:</td>
                  <td><%= session("pallets4") %></td>
                </tr>
                <tr>
                  <td>L / W / H:</td>
                  <td><%= session("length4") %> cm / <%= session("width4") %> cm / <%= session("height4") %> cm</td>
                </tr>
                <tr>
                  <td>&nbsp;</td>
                  <td><%= session("weight4") %> kg</td>
                </tr>
              </table>
              <% end if %>
              <% if len(session("description6")) > 0 or len(session("item_ref6")) > 0 or len(session("items6")) > 0 or len(session("pallets6")) > 0 or len(session("length6")) > 0 or len(session("width6")) > 0 or len(session("height6")) > 0 or len(session("weight6")) > 0 then %>
              <br />
              <table width="450" cellpadding="3" cellspacing="0" class="thin_border">
                <tr>
                  <td colspan="2" class="item_maintenance_header">Line 6</td>
                </tr>
                <tr>
                  <td width="30%">Description:</td>
                  <td width="70%"><%= session("description6") %></td>
                </tr>
                <tr>
                  <td>Ref:</td>
                  <td><%= session("item_ref6") %></td>
                </tr>
                <tr>
                  <td>Qty:</td>
                  <td><%= session("items6") %></td>
                </tr>
                <tr>
                  <td>Pallets:</td>
                  <td><%= session("pallets6") %></td>
                </tr>
                <tr>
                  <td>L / W / H:</td>
                  <td><%= session("length6") %> cm / <%= session("width6") %> cm / <%= session("height6") %> cm</td>
                </tr>
                <tr>
                  <td>&nbsp;</td>
                  <td><%= session("weight6") %> kg</td>
                </tr>
              </table>
              <% end if %>
              <% if len(session("description8")) > 0 or len(session("item_ref8")) > 0 or len(session("items8")) > 0 or len(session("pallets8")) > 0 or len(session("length8")) > 0 or len(session("width8")) > 0 or len(session("height8")) > 0 or len(session("weight8")) > 0 then %>
              <br />
              <table width="450" cellpadding="3" cellspacing="0" class="thin_border">
                <tr>
                  <td colspan="2" class="item_maintenance_header">Line 8</td>
                </tr>
                <tr>
                  <td width="30%">Description:</td>
                  <td width="70%"><%= session("description8") %></td>
                </tr>
                <tr>
                  <td>Ref:</td>
                  <td><%= session("item_ref8") %></td>
                </tr>
                <tr>
                  <td>Qty:</td>
                  <td><%= session("items8") %></td>
                </tr>
                <tr>
                  <td>Pallets:</td>
                  <td><%= session("pallets8") %></td>
                </tr>
                <tr>
                  <td>L / W / H:</td>
                  <td><%= session("length8") %> cm / <%= session("width8") %> cm / <%= session("height8") %> cm</td>
                </tr>
                <tr>
                  <td>&nbsp;</td>
                  <td><%= session("weight8") %> kg</td>
                </tr>
              </table>
              <% end if %></td>
          </tr>
          <tr>
            <td><table width="450" cellpadding="3" cellspacing="0" border="0">
                <tr>
                  <td>Supplier:</td>
                  <td><select name="cboSupplier">
                      <option <% if session("supplier") = "" then Response.Write " selected" end if%> value="">...</option>
                      <option <% if session("supplier") = "Cope" then Response.Write " selected" end if%> value="Cope">Cope</option>
                      <option <% if session("supplier") = "StarTrack" then Response.Write " selected" end if%> value="StarTrack">StarTrack</option>
                      <option <% if session("supplier") = "Schenker" then Response.Write " selected" end if%> value="Schenker">Schenker</option>
                      <option <% if session("supplier") = "Kings" then Response.Write " selected" end if%> value="Kings">Kings</option>
                    </select></td>
                </tr>
                <tr>
                  <td width="30%">Con-note<span class="mandatory">*</span>:</td>
                  <td width="70%"><input type="text" id="txtConnote" name="txtConnote" maxlength="20" size="30" value="<%= session("consignment_no") %>" /></td>
                </tr>
                <tr>
                  <td>Comments:</td>
                  <td><textarea name="txtComments" id="txtComments" cols="40" rows="5"><%= session("comments") %></textarea></td>
                </tr>
                <tr>
                  <td>Status:</td>
                  <td><select name="cboStatus">
                      <option <% if session("status") = "1" then Response.Write " selected" end if%> value="1">Open</option>
                      <option <% if session("status") = "0" then Response.Write " selected" end if%> value="0">Completed</option>
                    </select>
                    <% if session("status") = "0" then %>
                    <img src="images/forward_arrow.gif" border="0" /> <a href="send-completed-email.asp">Notify requester upon completion</a>
                    <% end if %></td>
                </tr>
                <tr>
                  <td>Pick up?</td>
                  <td><select name="cboPickup">
                      <option <% if session("pickup") = "0" then Response.Write " selected" end if%> value="0">No</option>
                      <option <% if session("pickup") = "1" then Response.Write " selected" end if%> value="1">Yes</option>
                    </select>
                    <% if session("pickup") = "1" then %>
                    <img src="images/forward_arrow.gif" border="0" /> <a href="send-pickup-email.asp">Notify requester upon pickup</a>
                    <% end if %></td>
                </tr>
              </table></td>
            <td valign="bottom">&nbsp;</td>
          </tr>          
        </table>
        <p><input type="hidden" name="Action" />
              <input type="submit" value="Update Freight" /></p>
      </form>
      <h2>Comments<br />
        <img src="images/comment_bar.jpg" border="0" /></h2>
      <table cellpadding="5" cellspacing="0" border="0" class="comments_box">        
        <%= strCommentsList %>
        <tr>
          <td><form action="" method="post" onsubmit="return submitComment(this)">
              <p><input type="text" name="txtComment" id="txtComment" maxlength="60" size="65" />
              <input type="hidden" name="Action" />
              <input type="submit" value="Add Comment" /></p>
            </form></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
</body>
</html>