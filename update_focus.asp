<% response.cookies("current_URL_cookie_logistics") = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?" & Request.Querystring %>
<!--#include file="include/connection_it.asp " -->
<% strSection = "roadshow" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Update FOCUS Item</title>
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="include/generic_form_validations.js"></script>
<script language="JavaScript" type="text/javascript">
function validateFormOnSubmit(theForm) {
	var reason = "";
	var blnSubmit = true;

	reason += validateEmptyField(theForm.txtProductCode,"Product Code");
	
	reason += validateNumeric(theForm.txtQuantity,"Quantity");

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

</script>
<%
Sub getItem
	dim intID
	intID = request("id")

    call OpenDataBase()

	set rs = Server.CreateObject("ADODB.recordset")

	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic

	strSQL = "SELECT * FROM tbl_focus WHERE id = " & intID

	rs.Open strSQL, conn

	'Response.Write strSQL

    if not DB_RecSetIsEmpty(rs) Then		
		session("owner") 			= rs("owner")
		session("type") 			= rs("type")		
		session("quantity") 		= rs("quantity")
		session("product_code") 	= rs("product_code")
		session("location") 		= rs("location")
		session("stock_situation") 	= rs("stock_situation")
		session("loan_account") 	= rs("loan_account")
		session("display") 			= rs("display")		
		session("available_for_sale") = rs("available_for_sale")
		session("notes") 			= rs("notes")
		session("instruction") 		= rs("instruction")
		session("comments") 		= rs("comments")		
		session("status") 			= rs("status")
		session("date_modified") 	= rs("date_modified")
		session("modified_by") 		= rs("modified_by")
		session("date_created") 	= rs("date_created")
		session("created_by") 		= rs("created_by")		
		session("logistics_action") = rs("logistics_action")
		session("invoice_no") 		= rs("invoice_no")
		session("pallet_no") 		= rs("pallet_no")
		session("loading_sequence") = rs("loading_sequence")
    end if

    call CloseDataBase()
end sub

sub updateItem
	dim strSQL
	dim intID
	intID = request("id")

	Call OpenDataBase()

	strSQL = "UPDATE tbl_focus SET "	
	strSQL = strSQL & "type = '" & Trim(Request.Form("cboType")) & "',"	
	strSQL = strSQL & "quantity = '" & Trim(Request.Form("txtQuantity")) & "',"
	strSQL = strSQL & "product_code = '" & Server.HTMLEncode(Trim(Request.Form("txtProductCode"))) & "',"
	strSQL = strSQL & "location = '" & Trim(Request.Form("cboLocation")) & "',"
	strSQL = strSQL & "stock_situation = '" & Trim(Request.Form("txtStockSituation")) & "',"
	strSQL = strSQL & "loan_account = '" & Trim(Request.Form("txtLoanAccount")) & "',"
	strSQL = strSQL & "display = '" & Trim(Request.Form("txtDisplay")) & "',"	
	strSQL = strSQL & "available_for_sale = '" & Trim(Request.Form("cboAvailableForSale")) & "',"
	strSQL = strSQL & "instruction = '" & Trim(Request.Form("txtInstruction")) & "',"
	strSQL = strSQL & "notes = '" & Replace(Request.Form("txtNotes"),"'","''") & "',"
	strSQL = strSQL & "logistics_action = '" & Trim(Request.Form("chkLogisticsAction")) & "',"
	strSQL = strSQL & "invoice_no = '" & Trim(Request.Form("txtInvoiceNo")) & "',"
	strSQL = strSQL & "pallet_no = '" & Server.HTMLEncode(Trim(Request.Form("txtPalletNo"))) & "',"
	strSQL = strSQL & "loading_sequence = '" & Server.HTMLEncode(Trim(Request.Form("txtLoadingSequence"))) & "',"	
	strSQL = strSQL & "date_modified = getdate(),"
	strSQL = strSQL & "modified_by = '" & session("UsrUserName") & "',"	
	strSQL = strSQL & "status = '" & Trim(Request.Form("cboStatus")) & "' WHERE id = " & intID

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

sub main
	call UTL_validateLogin
	
	if Request.ServerVariables("REQUEST_METHOD") = "POST" then
		if Trim(Request("Action")) = "Update" then
			call updateItem			
		end if
	end if
	
	call getItem
end sub

dim strMessageText
call main
%>
</head>
<body>
<table width="100%" cellpadding="0" cellspacing="0">
  <!-- #include file="include/header.asp" -->
  <tr>
    <td class="first_content"><table cellpadding="5" cellspacing="0" border="0">
        <tr>
          <td valign="top"><img src="images/backward_arrow.gif" border="0" /> <a href="list_focus.asp">Back to List</a>
            <h2>Update FOCUS item</h2>
            <font color="green"><%= strMessageText %></font></td>
          <td valign="top"><table cellpadding="4" cellspacing="0" class="created_table">
              <tr>
                <td class="created_column_1"><strong>Created:</strong></td>
                <td class="created_column_2"><%= session("created_by") %></td>
                <td class="created_column_3"><%= session("date_created") %></td>
              </tr>
              <tr>
                <td class="created_column_1"><strong>Last modified:</strong></td>
                <td class="created_column_2"><%= session("modified_by") %></td>
                <td class="created_column_3"><%= displayDateFormatted(session("date_modified")) %></td>
              </tr>
            </table></td>
        </tr>
      </table>
      <form action="" method="post" name="form_update_roadshow" id="form_update_roadshow" onsubmit="return validateFormOnSubmit(this)">
        <table border="0" cellpadding="5" cellspacing="0" width="1024">
          <tr>
            <td width="50%" valign="top"><table cellpadding="5" cellspacing="0" class="item_maintenance_box">
                <tr>
                  <td colspan="2" class="item_maintenance_header">Item Details</td>
                </tr>
                <tr>
                  <td align="right">Type:</td>
                  <td><select name="cboType">
                      <option <% if session("type") = "STOCK" then Response.Write " selected" end if%> value="STOCK">STOCK</option>
                      <option <% if session("type") = "POS" then Response.Write " selected" end if%> value="POS">POS</option>
                    </select></td>
                </tr>
                <tr>
                  <td width="25%" align="right">Product code<span class="mandatory">*</span>:</td>
                  <td width="75%"><input type="text" id="txtProductCode" name="txtProductCode" maxlength="20" size="20" value="<%= session("product_code") %>" /></td>
                </tr>
                <tr>
                  <td align="right">Qty<span class="mandatory">*</span>:</td>
                  <td><input type="text" id="txtQuantity" name="txtQuantity" maxlength="4" size="4" value="<%= session("quantity") %>" /></td>
                </tr>
                <tr>
                  <td align="right">Location:</td>
                  <td><select name="cboLocation">
                      <option <% if session("location") = "" then Response.Write " selected" end if%> value="">...</option>
                      <option <% if session("location") = "3T" then Response.Write " selected" end if%> value="3T">3T</option>
                      <option <% if session("location") = "HO" then Response.Write " selected" end if%> value="HO">HO</option>
                      <option <% if session("location") = "OTHER" then Response.Write " selected" end if%> value="OTHER">OTHER</option>
                    </select></td>
                </tr>
                <tr>
                  <td align="right">Stock Situation<span class="mandatory">*</span>:</td>
                  <td><input type="text" id="txtStockSituation" name="txtStockSituation" maxlength="60" size="30" value="<%= session("stock_situation") %>" /></td>
                </tr>
                <tr>
                  <td align="right">Loan Account:</td>
                  <td><input type="text" id="txtLoanAccount" name="txtLoanAccount" maxlength="15" size="15" value="<%= session("loan_account") %>" /></td>
                </tr>
                <tr>
                  <td align="right">Display:</td>
                  <td><input type="text" id="txtDisplay" name="txtDisplay" maxlength="60" size="30" value="<%= session("display") %>" /></td>
                </tr>
                <tr>
                  <td align="right">For sale?</td>
                  <td><select name="cboAvailableForSale">
                      <option <% if session("available_for_sale") = "NO" then Response.Write " selected" end if%> value="NO">No</option>
                      <option <% if session("available_for_sale") = "YES" then Response.Write " selected" end if%> value="YES">Yes</option>
                    </select></td>
                </tr>
                <tr>
                  <td align="right">Instruction:</td>
                  <td><input type="text" id="txtInstruction" name="txtInstruction" maxlength="60" size="30" value="<%= session("instruction") %>" /></td>
                </tr>
                <tr>
                  <td align="right">Notes:</td>
                  <td>
                  <input type="text" id="txtNotes" name="txtNotes" maxlength="90" size="40" value="<%= session("notes") %>" />
                  </td>
                </tr>
              </table></td>
            <td width="50%" valign="top"><table cellpadding="5" cellspacing="0" class="item_maintenance_box">
                <tr align="right" bgcolor="#99FF66">
                  <td colspan="2"><input type="checkbox" name="chkLogisticsAction" id="chkLogisticsAction" value="1" <% if session("logistics_action") = "1" then Response.Write " checked" end if%> /><label for="chkLogisticsAction">
                    Logistics Actioned</label></td>
                </tr>
                <tr class="status_row">
                  <td width="30%" align="right">Invoice no:</td>
                  <td width="70%"><input type="text" id="txtInvoiceNo" name="txtInvoiceNo" maxlength="20" size="20" value="<%= session("invoice_no") %>" /></td>
                </tr>
                <tr class="status_row">
                  <td align="right">Pallet no:</td>
                  <td><input type="text" id="txtPalletNo" name="txtPalletNo" maxlength="20" size="20" value="<%= session("pallet_no") %>" /></td>
                </tr>
                <tr class="status_row">
                  <td align="right">Loading sequence:</td>
                  <td><input type="text" id="txtLoadingSequence" name="txtLoadingSequence" maxlength="20" size="20" value="<%= session("loading_sequence") %>" /></td>
                </tr>
                <tr>
                  <td align="right">Status:</td>
                  <td><select name="cboStatus">
                      <option <% if session("status") = "1" then Response.Write " selected" end if%> value="1">Open</option>
                      <option <% if session("status") = "0" then Response.Write " selected" end if%> value="0">Completed</option>
                    </select></td>
                </tr>
              </table></td>
          </tr>
          <tr>
            <td colspan="2" valign="top" align="center"><input type="hidden" name="Action" />
              <input type="submit" value="Update" /></td>
          </tr>
        </table>
      </form></td>
  </tr>
</table>
</body>
</html>