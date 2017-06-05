<!--#include file="include/connection_it.asp " -->
<!--#include file="class/clsComment.asp " -->
<% strSection = "cancelled" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Update Cancelled Order</title>
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="include/usableforms.js"></script>
<script type="text/javascript" src="include/generic_form_validations.js"></script>
<script language="JavaScript" type="text/javascript">
function validateFormOnSubmit(theForm) {
	var reason 		= "";
	var blnSubmit 	= true;
	
	reason += validateEmptyField(theForm.txtShipmentNo,"Shipment no");
	reason += validateEmptyField(theForm.txtInfo,"Info");	
	
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

function cancelOrder(theForm) {
	if (confirm ("Please click OK to confirm this order cancellation.")){
		theForm.Action.value = 'Cancel';
		return true;
    } else {
		return false;
	}
}

function completeOrder(theForm) {
	if (confirm ("Please click OK to complete this order cancellation")){
		theForm.Action.value = 'Complete';
		return true;
    } else {
		return false;
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
Sub getCancelledOrder	
	dim strSQL
	
    call OpenDataBase()

	set rs = Server.CreateObject("ADODB.recordset")

	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic

	strSQL = "SELECT * FROM yma_cancelled_order WHERE cancel_id = " & session("cancel_id")

	rs.Open strSQL, conn

    if not DB_RecSetIsEmpty(rs) Then
		session("cancel_shipment_no") 				= rs("cancel_shipment_no")
		session("cancel_info") 						= rs("cancel_info")
		session("cancel_comments") 					= rs("cancel_comments")
		session("cancel_warehouse_confirm") 		= rs("cancel_warehouse_confirm")
		session("cancel_warehouse_confirm_by") 		= rs("cancel_warehouse_confirm_by")
		session("cancel_warehouse_confirm_date") 	= rs("cancel_warehouse_confirm_date")
		session("cancel_logistics_confirm") 		= rs("cancel_logistics_confirm")
		session("cancel_logistics_confirm_by") 		= rs("cancel_logistics_confirm_by")
		session("cancel_logistics_confirm_date") 	= rs("cancel_logistics_confirm_date")
		session("cancel_created_by") 				= rs("cancel_created_by")
		session("cancel_date_created") 				= rs("cancel_date_created")
		session("cancel_modified_by") 				= rs("cancel_modified_by")
		session("cancel_date_modified") 			= rs("cancel_date_modified")
		session("cancel_status") 					= rs("cancel_status")
    end if

    call CloseDataBase()
end sub

sub updateCancelledOrder
	dim strSQL
	
	dim strShipmentNo
	dim strInfo
	dim strComments
	
	strShipmentNo 	= Trim(Request.Form("txtShipmentNo"))
	strInfo 		= Trim(Request.Form("txtInfo"))
	strComments 	= Replace(Request.Form("txtComments"),"'","''")
	
	call OpenDataBase()
	
	strSQL = "UPDATE yma_cancelled_order SET "
	strSQL = strSQL & " cancel_shipment_no = '" & Server.HTMLEncode(strShipmentNo) & "', "
	strSQL = strSQL & " cancel_info = '" & Server.HTMLEncode(strInfo) & "', "
	strSQL = strSQL & " cancel_comments = '" & Server.HTMLEncode(strComments) & "', "
	strSQL = strSQL & " cancel_date_modified = getdate(), "
	strSQL = strSQL & " cancel_modified_by = '" & session("UsrUserName") & "' "
	strSQL = strSQL & " WHERE cancel_id = " & session("cancel_id")
	
	'response.Write strSQL
	  
	on error resume next
	conn.Execute strSQL
	
	On error Goto 0
	
	if err <> 0 then
		strMessageText = err.description
	else		
		strMessageText = "The record has been updated."
	end if 
	
	call CloseDataBase()
end sub

sub updateWarehouseConfirm
	dim strSQL
	
	call OpenDataBase()
	
	strSQL = "UPDATE yma_cancelled_order SET "
	strSQL = strSQL & " cancel_warehouse_confirm = '1', "
	strSQL = strSQL & " cancel_warehouse_confirm_by = '" & session("UsrUserName") & "', "
	strSQL = strSQL & " cancel_warehouse_confirm_date = getdate(), "
	strSQL = strSQL & " cancel_date_modified = getdate(), "
	strSQL = strSQL & " cancel_modified_by = '" & session("UsrUserName") & "' "
	strSQL = strSQL & " WHERE cancel_id = " & session("cancel_id")
	
	'response.Write strSQL
	Set oMail = Server.CreateObject("CDO.Message")
	Set iConf = Server.CreateObject("CDO.Configuration")
	Set Flds = iConf.Fields
				
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.sendgrid.net"
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 10
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 'basic clear text
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "yamahamusicau"
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "str0ppy@16"
	iConf.Fields.Update
				
	emailFrom 		= "automailer@music.yamaha.com"
	  
	on error resume next
	conn.Execute strSQL
	
	On error Goto 0
	
	if err <> 0 then
		strMessageText = err.description
	else
		emailTo 	= "logistics-aus@music.yamaha.com"
		emailCc 	= session("cancel_created_by")		
		emailSubject = "Warehouse confirmed Cancelled Order (Ship no: " & session("cancel_shipment_no") & ") by " & session("UsrUserName")
		emailBodyText =	"Requested by : " & session("cancel_created_by")  & vbCrLf _											
					&	"---------------------------------------------------------------------------" & vbCrLf _				
					&	"Shipment no  : " & session("cancel_shipment_no") & vbCrLf _
					&	"Info         : " & session("cancel_info") & vbCrLf _
					&	"---------------------------------------------------------------------------" & vbCrLf _
					&	"Warehouse confirmed by : " & session("UsrUserName") & vbCrLf _
					&	"---------------------------------------------------------------------------" & vbCrLf _																	
					&	" " & vbCrLf _
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
		
		strMessageText = "The record has been updated."
	end if 
	
	call CloseDataBase()
end sub

sub updateLogisticsConfirm
	dim strSQL
	
	call OpenDataBase()
	
	strSQL = "UPDATE yma_cancelled_order SET "
	strSQL = strSQL & " cancel_status = '0', "
	strSQL = strSQL & " cancel_logistics_confirm = '1', "
	strSQL = strSQL & " cancel_logistics_confirm_by = '" & session("UsrUserName") & "', "
	strSQL = strSQL & " cancel_logistics_confirm_date = getdate(), "
	strSQL = strSQL & " cancel_date_modified = getdate(), "
	strSQL = strSQL & " cancel_modified_by = '" & session("UsrUserName") & "' "
	strSQL = strSQL & " WHERE cancel_id = " & session("cancel_id")
	
	'response.Write strSQL
	Set oMail = Server.CreateObject("CDO.Message")
	Set iConf = Server.CreateObject("CDO.Configuration")
	Set Flds = iConf.Fields
				
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.sendgrid.net"
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 10
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 'basic clear text
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "yamahamusicau"
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "str0ppy@16"
	iConf.Fields.Update
				
	emailFrom 		= "logistics-aus@music.yamaha.com"
	  
	on error resume next
	conn.Execute strSQL
	
	On error Goto 0
	
	if err <> 0 then
		strMessageText = err.description
	else
		emailTo 		= session("cancel_created_by")
		emailCc			= "YMA-Warehouse@ttlogistics.com.au"
		'emailBcc 		= "logistics-aus@music.yamaha.com"
		emailSubject 	= "Logistics completed the Cancelled Order (Ship no: " & session("cancel_shipment_no") & ") by " & session("UsrUserName")
		emailBodyText 	= "Requested by : " & session("cancel_created_by")  & vbCrLf _											
					&	"---------------------------------------------------------------------------" & vbCrLf _
					&	"Shipment no  : " & session("cancel_shipment_no") & vbCrLf _
					&	"Info         : " & session("cancel_info") & vbCrLf _
					&	"---------------------------------------------------------------------------" & vbCrLf _
					&	"Completed by : " & session("UsrUserName") & vbCrLf _
					&	"---------------------------------------------------------------------------" & vbCrLf _
					&	" " & vbCrLf _
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
		
		strMessageText = "The record has been updated."
	end if 
	
	call CloseDataBase()
end sub

sub main
	call UTL_validateLogin
	
	session("cancel_id") = request("id")	
	
	if Request.ServerVariables("REQUEST_METHOD") = "POST" then
		select case Trim(Request("Action"))
			case "Update"
				call updateCancelledOrder				
			case "Cancel"
				call updateWarehouseConfirm				
			case "Complete"
				call updateLogisticsConfirm				
			case "Comment"
				call addComment(session("cancel_id"),cancelledModuleID)				
		end select
	end if
	
	call getCancelledOrder
	call listComments(session("cancel_id"),cancelledModuleID)
end sub

call main

dim strDamageTypeList
dim strCourseDamageList
dim strCommentsList
%>
</head>
<body>
<table cellspacing="0" cellpadding="0" align="center" class="full_size_table" border="0">
  <!-- #include file="include/header.asp" -->
  <tr>
    <td class="first_content"><table cellpadding="5" cellspacing="0" border="0">
        <tr>
          <td><a href="list_cancelled.asp"><img src="images/icon_cancelled.jpg" border="0" alt="Cancelled Order" /></a></td>
          <td valign="top"><img src="images/backward_arrow.gif" border="0" /> <a href="list_cancelled.asp">Back to List</a>
            <h2>Update Cancelled Order</h2>
            <font color="green"><%= strMessageText %></font></td>
          <td valign="top"><table cellpadding="4" cellspacing="0" class="created_table_wide">
              <tr>
                <td class="created_column_1"><strong>Created:</strong></td>
                <td class="created_column_2"><%= Lcase(session("cancel_created_by")) %></td>
                <td class="created_column_3"><%= displayDateFormatted(session("cancel_date_created")) %></td>
              </tr>
              <tr>
                <td><strong>Last modified:</strong></td>
                <td><%= session("cancel_modified_by") %></td>
                <td><%= displayDateFormatted(session("cancel_date_modified")) %></td>
              </tr>
            </table></td>
        </tr>
      </table>
      <table border="0" cellpadding="5" cellspacing="0" width="1024">
        <tr>
          <td width="50%" valign="top"><form action="" method="post" name="form_add_cancelled_order" id="form_add_cancelled_order" onsubmit="return validateFormOnSubmit(this)">
              <table cellpadding="5" cellspacing="0" class="item_maintenance_box">
                <tr>
                  <td class="item_maintenance_header">Order Details</td>
                </tr>
                <tr>
                  <td>Shipment no<span class="mandatory">*</span>:<br />
                    <input type="text" id="txtShipmentNo" name="txtShipmentNo" maxlength="20" size="30" value="<%= session("cancel_shipment_no") %>" /></td>
                </tr>
                <tr>
                  <td>Reason for cancellation<span class="mandatory">*</span>:<br />
                    <input type="text" id="txtInfo" name="txtInfo" maxlength="70" size="75" value="<%= session("cancel_info") %>" /></td>
                </tr>
                <tr>
                  <td>Comments:<br />
              <textarea name="txtComments" id="txtComments" cols="55" rows="5"><%= session("cancel_comments") %></textarea></td>
                </tr>
                <tr class="status_row">
                  <td>Status: <u>
                  <% select case session("cancel_status")
				  		case 1
							response.Write("Open")
						case 2
							response.Write("Cancelled")
						case 0
							response.Write("Completed")
					end select
				  %>
                    </u></td>
                </tr>
                <tr>
                  <td><input type="hidden" name="Action" />
                    <input type="submit" value="Update Cancelled Order" /></td>
                </tr>
              </table>
          </form></td>
          <td width="50%" valign="top"><table cellpadding="5" cellspacing="0" class="item_maintenance_box">
              <tr>
                <td class="item_maintenance_header">1. Warehouse confirm</td>
              </tr>
              <tr>
                <td align="center">
				<% if session("cancel_warehouse_confirm") = 0 then %>
                  <form action="" method="post" name="form_warehouse_cancel" id="form_warehouse_cancel" onsubmit="return cancelOrder(this)">
                    <input type="hidden" name="Action" />
                    <input type="submit" value="Cancel Order" />
                  </form>
                  <% else %>
                  <%= session("cancel_warehouse_confirm_by") %> - 
                  <%= displayDateFormatted(session("cancel_warehouse_confirm_date")) %>
                  <% end if %></td>
              </tr>
            </table>
            <br />
            <table cellpadding="5" cellspacing="0" class="item_maintenance_box">
              <tr>
                <td class="item_maintenance_header">2. Logistics confirm</td>
              </tr>
              <tr>
                <td align="center"><% if session("cancel_warehouse_confirm") = 1 and session("cancel_logistics_confirm") = 0 then %>
                  <form action="" method="post" name="form_logistics_cancel" id="form_logistics_cancel" onsubmit="return completeOrder(this)">
                    <input type="hidden" name="Action" />
                    <input type="submit" value="Complete" />
                  </form>
                  <% else %>
                        <%= session("cancel_logistics_confirm_by") %> - 
                        <%= displayDateFormatted(session("cancel_logistics_confirm_date")) %>
                  <% end if %></td>
              </tr>
            </table></td>
        </tr>
      </table>
      <h2>Comments<br />
        <img src="images/comment_bar.jpg" border="0" /></h2>
      <table cellpadding="5" cellspacing="0" border="0" class="comments_box">
        <%= strCommentsList %>
        <tr>
          <td><form action="" name="form_add_comment" id="form_add_comment" method="post" onsubmit="return submitComment(this)">
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
</body>
</html>