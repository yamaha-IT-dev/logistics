<% response.cookies("current_URL_cookie_logistics") = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?" & Request.Querystring %>
<!--#include file="include/connection_it.asp " -->
<!--#include file="class/clsComment.asp " -->
<!--#include file="class/clsUnsolicited.asp " -->
<% strSection = "unsolicited" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Update Unsolicited Goods</title>
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="include/generic_form_validations.js"></script>
<script language="JavaScript" type="text/javascript">
function validateFormOnSubmit(theForm) {
	var reason 		= "";
	var blnSubmit 	= true;
	
	reason += validateEmptyField(theForm.txtItemCode,"Item Code");
	reason += validateSpecialCharacters(theForm.txtItemCode,"Item Code");
	
	reason += validateEmptyField(theForm.txtDescription,"Description");
	reason += validateSpecialCharacters(theForm.txtDescription,"Description");
	
	reason += validateSpecialCharacters(theForm.txtDealer,"Dealer");
	
	reason += validateEmptyField(theForm.txtShipmentNo,"Shipment No");
	reason += validateSpecialCharacters(theForm.txtShipmentNo,"Shipment No");
	
	reason += validateNumeric(theForm.txtQty,"Qty");
	
	reason += validateEmptyField(theForm.txtConnote,"Return Connote");
	reason += validateSpecialCharacters(theForm.txtConnote,"Return Connote");

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
sub main
	dim unsID
	unsID = request("id")
	
	if Request.ServerVariables("REQUEST_METHOD") = "POST" then
		select case Trim(Request("Action"))
			case "Update"
				unsID 			= Request("id")
				unsDepartment 	= Request.Form("cboDepartment")
				unsItemCode 	= Request.Form("txtItemCode")
				unsDescription 	= Request.Form("txtDescription")
				unsConnote 		= Request.Form("txtConnote")
				unsGRA 			= Request.Form("txtGRA")
				unsDealer 		= Request.Form("txtDealer")
				unsShipmentNo 	= Request.Form("txtShipmentNo")
				unsQty 			= Request.Form("txtQty")
				unsInstruction 	= Request.Form("cboInstruction")
				unsComments 	= Request.Form("txtComments")
				unsStatus		= Request.Form("cboStatus")
				call updateUnsolicited(unsID, unsDepartment, unsItemCode, unsDescription, unsConnote, unsGRA, unsDealer, unsShipmentNo, unsQty, unsInstruction, unsComments, unsStatus, session("UsrUserName"))	
			case "Comment"
				'unsID 			= Request("id")
				call addComment(unsID,unsolicitedGoodsModuleID)
		end select
	end if
	
	call UTL_validateLogin
		
	
		
	call getUnsolicited(unsID)
	call listComments(unsID,unsolicitedGoodsModuleID)
end sub

call main

dim strMessageText
dim strCommentsList

Dim unsDepartment, unsItemCode, unsDescription, unsConnote, unsGRA, unsDealer, unsShipmentNo, unsQty, unsInstruction, unsComments, unsStatus, unsCreatedBy
%>
</head>
<body>
<table cellspacing="0" cellpadding="0" align="center" class="full_size_table" border="0">
  <!-- #include file="include/header.asp" -->
  <tr>
    <td class="first_content"><table cellpadding="5" cellspacing="0" border="0">
        <tr>
          <td><a href="list_quarantine.asp"><img src="images/icon_return.jpg" border="0" alt="Warehouse Return" /></a></td>
          <td valign="top"><img src="images/backward_arrow.gif" border="0" /> <a href="list_unsolicited.asp">Back to List</a>
            <h2>Update Unsolicited Goods</h2>
            <font color="green"><%= strMessageText %></font></td>
          <td valign="top"><table cellpadding="4" cellspacing="0" class="created_table">
              <tr>
                <td class="created_column_1"><strong>Created:</strong></td>
                <td class="created_column_2"><%= session("unsCreatedBy") %></td>
                <td class="created_column_3"><%= displayDateFormatted(session("unsDateCreated")) %></td>
              </tr>
              <tr>
                <td><strong>Last modified:</strong></td>
                <td><%= session("unsModifiedBy") %></td>
                <td><%= displayDateFormatted(session("unsDateModified")) %></td>
              </tr>
            </table></td>
        </tr>
      </table>
      <form action="" method="post" name="form_update_unsolicited" id="form_update_unsolicited" onsubmit="return validateFormOnSubmit(this)">
        <table cellpadding="5" cellspacing="0" class="item_maintenance_box">
          <tr>
            <td colspan="3" class="item_maintenance_header">Unsolicited Goods to Excel</td>
          </tr>
          <tr>
            <td>Department:<br />
              <select name="cboDepartment">
                <option <% if session("unsDepartment") = "AV" then Response.Write " selected" end if%> value="AV">AV</option>
                <option <% if session("unsDepartment") = "MPD" then Response.Write " selected" end if%> value="MPD">MPD</option>
                <option <% if session("unsDepartment") = "Other" then Response.Write " selected" end if%> value="Other">Other</option>
            </select></td>
            <td colspan="2">Item code<span class="mandatory">*</span>:<br />
              <input type="text" id="txtItemCode" name="txtItemCode" maxlength="20" size="25" value="<%= Server.HTMLEncode(session("unsItemCode")) %>" /></td>
          </tr>
          <tr>
            <td colspan="3">Description<span class="mandatory">*</span>:<br />
              <input type="text" id="txtDescription" name="txtDescription" maxlength="50" size="60" value="<%= Server.HTMLEncode(session("unsDescription")) %>" /></td>
          </tr>
          <tr>
            <td>Con-note<span class="mandatory">*</span>:<br />
              <input type="text" id="txtConnote" name="txtConnote" maxlength="15" size="20" value="<%= Server.HTMLEncode(session("unsConnote")) %>" /></td>
            <td colspan="2">GRA:<br />
            <input type="text" id="txtGRA" name="txtGRA" maxlength="7" size="10" value="<%= Server.HTMLEncode(session("unsGRA")) %>" /></td>
          </tr>
          <tr>
            <td width="60%">Dealer:<br />
              <input type="text" id="txtDealer" name="txtDealer" maxlength="30" size="35" value="<%= Server.HTMLEncode(session("unsDealer")) %>" /></td>
            <td width="20%">Shipment no<span class="mandatory">*</span>:<br />
              <input type="text" id="txtShipmentNo" name="txtShipmentNo" maxlength="10" size="12" value="<%= Server.HTMLEncode(session("unsShipmentNo")) %>" /></td>
            <td width="20%">Qty<span class="mandatory">*</span>:<br />
              <input type="text" id="txtQty" name="txtQty" maxlength="4" size="5" value="<%= Server.HTMLEncode(session("unsQty")) %>" /></td>
          </tr>
          <tr>
            <td>Instruction:<br />
              <select name="cboInstruction">
                <option value="1" <% if session("unsInstruction") = "1" then Response.Write " selected" end if %> >Move to 3XL</option>
                <option value="2" <% if session("unsInstruction") = "2" then Response.Write " selected" end if %> >Move to 3S</option>
                <option value="3" <% if session("unsInstruction") = "3" then Response.Write " selected" end if %> >Investigate</option>
              </select></td>
            <td colspan="2">&nbsp;</td>
          </tr>
          <tr>
            <td colspan="3">Comments:<br />
              <textarea name="txtComments" id="txtComments" cols="50" rows="3"><%= session("unsComments") %></textarea></td>
          </tr>
          <tr class="status_row">
            <td colspan="3">Status:
              <select name="cboStatus">
                <option <% if session("unsStatus") = "1" then Response.Write " selected" end if%> value="1">Open</option>
                <option <% if session("unsStatus") = "0" then Response.Write " selected" end if%> value="0">Completed</option>
              </select></td>
          </tr>
          <tr>
            <td colspan="3"><input type="hidden" name="Action" />
              <input type="submit" value="Update Unsolicited Goods" />
              <input type="reset" value="Reset" /></td>
          </tr>
        </table>
      </form>
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