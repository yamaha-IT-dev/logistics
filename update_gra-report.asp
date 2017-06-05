<% response.cookies("current_URL_cookie_logistics") = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?" & Request.Querystring %>
<!--#include file="include/connection_it.asp " -->
<!--#include file="class/clsGoodsReturn.asp " -->
<!--#include file="class/clsGoodsReturnReport.asp " -->
<!--#include file="class/clsPallet.asp " -->
<!--#include file="class/clsWarrantyCode.asp " -->
<% strSection = "gra" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Update GRA Report</title>
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="include/autoSum.js"></script>
<script type="text/javascript" src="include/usableforms.js"></script>
<script type="text/javascript" src="include/generic_form_validations.js"></script>
<script language="JavaScript" type="text/javascript">
function validateUpdateFormOnSubmit(theForm) {
	var reason = "";
	var blnSubmit = true;
	
	reason += validateEmptyField(theForm.txtGraNo,"GRA no");
	reason += validateSpecialCharacters(theForm.txtGraNo,"GRA no");
	
	reason += validateEmptyField(theForm.txtLine,"Line");
	
	reason += validateEmptyField(theForm.txtItem,"Item");
	reason += validateSpecialCharacters(theForm.txtItem,"Item");
	
	reason += validateEmptyField(theForm.txtSerialNo,"Serial no");
	reason += validateSpecialCharacters(theForm.txtSerialNo,"Serial no");
	
	reason += validateEmptyField(theForm.txtDealerCode,"Dealer code");
	reason += validateSpecialCharacters(theForm.txtDealerCode,"Dealer code");
	
	reason += validateEmptyField(theForm.cboWarrantyCode,"Warranty code");
	
	reason += validateEmptyField(theForm.txtRepairReport,"Repair Report");
	reason += validateSpecialCharacters(theForm.txtRepairReport,"Repair Report");
	
	reason += validateNumeric(theForm.txtLabour,"Labour");
	reason += validateSpecialCharacters(theForm.txtLabour,"Labour");
	
	reason += validateNumeric(theForm.txtParts,"Parts");
	reason += validateSpecialCharacters(theForm.txtParts,"Parts");
	
	if (theForm.cboDestination.value == "Destroy") {
		reason += validateEmptyField(theForm.cboPalletNo,"Pallet");
	}
	
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
function submitExport(theForm) {
	var blnSubmit = true;

	if (blnSubmit == true){
        theForm.Action.value = 'Export';
		
		return true;
    }
}

</script>
<%
'----------------------------------------------------------------------------------------
' Update GRA Report record
'----------------------------------------------------------------------------------------
sub updateGraReport
	dim strSQL
	
	dim intID
	intID = request("id")
	
	dim intGraNo
	dim intLine	
	dim strItem
	dim strSerialNo
	dim strDealerCode
	dim strWarrantyCode
	dim strRepairReport
	dim intLabour
	dim intParts
	dim intGST
	dim intTotalCost
	dim strDateReceived
	dim strDateRepaired
	dim strDestination
	dim strPalletNo
	dim strComments
	dim intStatus
	
	intGraNo 			= Trim(Request.Form("txtGraNo"))
	intLine 			= Trim(Request.Form("txtLine"))
	strItem 			= Replace(Request.Form("txtItem"),"'","''")
	strSerialNo			= Replace(Request.Form("txtSerialNo"),"'","''")
	strDealerCode		= Replace(Request.Form("txtDealerCode"),"'","''")
	strWarrantyCode		= Trim(Request.Form("cboWarrantyCode"))
	strRepairReport 	= Replace(Request.Form("txtRepairReport"),"'","''")
	intLabour			= Trim(Request.Form("txtLabour"))
	intParts			= Trim(Request.Form("txtParts"))
	intGST				= Trim(Request.Form("txtGST"))
	intTotalCost		= Trim(Request.Form("txtTotalCost"))
	strDateReceived		= Trim(Request.Form("txtDateReceived"))
	strDateRepaired		= Trim(Request.Form("txtDateRepaired"))
	strDestination		= Trim(Request.Form("cboDestination"))
	strPalletNo			= Trim(Request.Form("cboPalletNo"))
	strComments 		= Replace(Request.Form("txtComments"),"'","''")
	intStatus			= Trim(Request.Form("cboStatus"))
	
	Call OpenDataBase()

	strSQL = "UPDATE yma_gra_report SET "
	strSQL = strSQL & "gra_no = '" & Server.HTMLEncode(intGraNo) & "',"
	strSQL = strSQL & "line_no = '" & intLine & "',"
	strSQL = strSQL & "item = '" & Server.HTMLEncode(strItem) & "',"
	strSQL = strSQL & "serial_no = '" & Server.HTMLEncode(strSerialNo) & "',"
	strSQL = strSQL & "dealer_code = '" & Server.HTMLEncode(strDealerCode) & "',"
	strSQL = strSQL & "gra_warranty_code = '" & strWarrantyCode & "',"
	strSQL = strSQL & "repair_report = '" & Server.HTMLEncode(strRepairReport) & "',"
	strSQL = strSQL & "labour = CONVERT(money," & intLabour & "),"
	strSQL = strSQL & "parts = CONVERT(money," & intParts & "),"
	strSQL = strSQL & "gst = CONVERT(money," & intGST & "),"
	strSQL = strSQL & "total = CONVERT(money," & intTotalCost & "),"
	strSQL = strSQL & "date_received = CONVERT(datetime,'" & strDateReceived & "',103),"
	strSQL = strSQL & "date_repaired = CONVERT(datetime,'" & strDateRepaired & "',103),"
	strSQL = strSQL & "date_modified = getdate(),"
	strSQL = strSQL & "modified_by = '" & session("UsrUserName") & "',"
	strSQL = strSQL & "destination = '" & strDestination & "',"
	strSQL = strSQL & "pallet_no = '" & Server.HTMLEncode(strPalletNo) & "',"
	strSQL = strSQL & "comments = '" & Server.HTMLEncode(strComments) & "',"
	strSQL = strSQL & "status = '" & intStatus & "' WHERE report_id = " & intID

	'response.Write strSQL
	
	on error resume next
	conn.Execute strSQL

	if err <> 0 then
		strMessageText = err.description
	else
		strMessageText = "The GRA report has been updated."
	end if

	Call CloseDataBase()
end sub

'----------------------------------------------------------------------------------------
' SET GRA REPORT INVOICE_EXPORTED FLAG TO 1
'----------------------------------------------------------------------------------------
sub updateGraReportExportedFlag
	dim strSQL
		
	Call OpenDataBase()

	strSQL = "UPDATE yma_gra_report SET "
	strSQL = strSQL & "invoice_exported = '1',"
	strSQL = strSQL & "date_exported = getdate(),"
	strSQL = strSQL & "status = '4' WHERE report_id = " & session("report_id")

	'response.Write strSQL
	
	on error resume next
	conn.Execute strSQL

	if err <> 0 then
		strMessageText = err.description
	else
		strMessageText = "The GRA report has been successfully exported."
	end if

	Call CloseDataBase()
end sub

sub main
	call UTL_validateLogin
	
	select case Trim(Request("ref"))
		case "gra"
			strBackLink = "list_gra.asp"
		case "report"
			strBackLink = "list_gra_report.asp"
		case else
			strBackLink = "list_gra.asp"	
	end select
	
	session("report_id")  	= Request("id")
	
	call getGraReport(session("report_id"))	
	call getGraFromBASE(session("report_gra_no"))
	call getGraStatus(session("report_gra_no"))
	call getPalletList
	
	if Request.ServerVariables("REQUEST_METHOD") = "POST" then	
		Select Case Trim(Request("Action"))
			case "Update"
				call updateGraReport
				call getGraReport(session("report_id"))
				call getPalletList
				call getWarrantyCodeList
		end select
	else
		call getWarrantyCodeList
	end if
end sub

dim strMessageText
dim strBackLink

call main

dim rs
dim intPageCount, intpage, intRecord
dim strDisplayList
dim strPalletList
dim strWarrantyCodeList
%>
</head>
<body>
<table width="100%" cellpadding="0" cellspacing="0">
  <!-- #include file="include/header.asp" -->
  <tr>
    <td class="first_content"><table cellpadding="5" cellspacing="0" border="0">
        <tr>
          <td valign="top"><a href="list_gra.asp"><img src="images/icon_gra.jpg" border="0" alt="GRA" /></a></td>
          <td valign="top" width="300"><img src="images/backward_arrow.gif" border="0" /> <a href="list_gra_report.asp">Back to List</a>
            <h2>Update GRA Report</h2>
            <font color="green"><%= strMessageText %></font></td>
          <td valign="top"><table cellpadding="4" cellspacing="0" class="created_table">
              <tr>
                <td width="30%"><strong>Report created by:</strong></td>
                <td width="20%"><%= session("report_created_by") %></td>
                <td width="50%"><%= displayDateFormatted(session("report_date_created")) %></td>
              </tr>
              <tr>
                <td><strong>Last modified:</strong></td>
                <td><%= session("report_modified_by") %></td>
                <td><%= displayDateFormatted(session("report_date_modified")) %></td>
              </tr>
            </table></td>
        </tr>
      </table>
      <form action="" method="post" name="form_gra" id="form_gra" onsubmit="return validateUpdateFormOnSubmit(this)">
        <table class="white_bordered_table_small" cellpadding="5" cellspacing="0">
          <tr>
            <td colspan="2" class="item_maintenance_header">Report Details</td>
          </tr>
          <tr>
            <td width="20%">Gra no<span class="mandatory">*</span>:</td>
            <td width="80%"><input type="text" id="txtGraNo" name="txtGraNo" maxlength="10" size="15" value="<%= session("report_gra_no") %>" />
            <font color="red">Don't enter the <strong>/XX</strong> in the end, it should be in the <strong>Line</strong></font></td>
          </tr>
          <tr>
            <td>Item<span class="mandatory">*</span>:</td>
            <td><input type="text" id="txtItem" name="txtItem" maxlength="40" size="50" value="<%= session("report_item") %>" /></td>
          </tr>
          <tr>
            <td>Line<span class="mandatory">*</span>:</td>
            <td><input type="text" id="txtLine" name="txtLine" maxlength="3" size="4" value="<%= session("report_line_no") %>" /></td>
          </tr>
          <tr>
            <td>Serial no<span class="mandatory">*</span>:</td>
            <td><input type="text" id="txtSerialNo" name="txtSerialNo" maxlength="20" size="30" value="<%= session("report_serial_no") %>" /></td>
          </tr>
          <tr>
            <td>Warranty code<span class="mandatory">*</span>:</td>
            <td><select name="cboWarrantyCode"><%= strWarrantyCodeList %></select></td>
          </tr>
          <tr>
            <td>Dealer code<span class="mandatory">*</span>:</td>
            <td><input type="text" id="txtDealerCode" name="txtDealerCode" maxlength="20" size="30" value="<%= session("report_dealer_code") %>" /></td>
          </tr>
          <tr>
            <td>Repair report<span class="mandatory">*</span>:</td>
            <td><input type="text" id="txtRepairReport" name="txtRepairReport" maxlength="90" size="70" value="<%= session("report_repair_report") %>" /></td>
          </tr>
          <tr>
            <td>Labour<span class="mandatory">*</span>:</td>
            <td>$
              <input type="text" id="txtLabour" name="txtLabour" maxlength="6" size="8" onfocus="startCalc();" onblur="stopCalc();"  value="<%= session("report_labour") %>" /></td>
          </tr>
          <tr>
            <td>Parts<span class="mandatory">*</span>:</td>
            <td>$
              <input type="text" id="txtParts" name="txtParts" maxlength="6" size="8" onfocus="startCalc();" onblur="stopCalc();"  value="<%= session("report_parts") %>" /></td>
          </tr>
          <tr>
            <td>GST 10%:</td>
            <td>$
              <input type="text" id="txtGST" name="txtGST" maxlength="6" size="8" value="<%= session("report_gst") %>" style="background-color:#CCC" />
              <em>Auto-generated</em></td>
          </tr>
          <tr>
            <td>Total:</td>
            <td>$
              <input type="text" id="txtTotalCost" name="txtTotalCost" maxlength="6" size="8" value="<%= session("report_total") %>" style="background-color:#CCC" />
              <em>Auto-generated</em></td>
          </tr>
          <tr>
            <td>Date received:</td>
            <td><input type="text" id="txtDateReceived" name="txtDateReceived" maxlength="10" size="10"  value="<%= session("report_date_received") %>" />
              <em>DD/MM/YYYY</em></td>
          </tr>
          <tr>
            <td>Date repaired:</td>
            <td><input type="text" id="txtDateRepaired" name="txtDateRepaired" maxlength="10" size="10"  value="<%= session("report_date_repaired") %>" />
              <em>DD/MM/YYYY</em></td>
          </tr>
          <tr>
            <td>Destination:</td>
            <td><select name="cboDestination">
                <option value="" <% if session("report_destination") = "" then Response.Write " selected" end if%> rel="none">...</option>
                <option value="3T" <% if session("report_destination") = "3T" then Response.Write " selected" end if%> rel="none">3T</option>
                <option value="3S" <% if session("report_destination") = "3S" then Response.Write " selected" end if%> rel="none">3S</option>
                <option value="Destroy" <% if session("report_destination") = "Destroy" then Response.Write " selected" end if%> rel="destroy">Destroy</option>
              </select></td>
          </tr>
          <tr rel="destroy">
            <td align="right">Write-off Pallet no:</td>
            <td><select name="cboPalletNo">
                <%= strPalletList %>
              </select></td>
          </tr>
          <tr>
            <td>Comments:</td>
            <td><textarea name="txtComments" id="txtComments" cols="50" rows="5" onKeyDown="limitText(this.form.txtComments,this.form.countdown,200);" 
onKeyUp="limitText(this.form.txtComments,this.form.countdown,200);"><%= session("report_comments") %></textarea></td>
          </tr>
          <tr class="status_row">
            <td>Report status:</td>
            <td><select name="cboStatus">
                <option <% if session("report_status") = "1" then Response.Write " selected" end if%> value="1" <% if session("report_status") = "0" then Response.Write " disabled" end if%>>Open</option>
                <option <% if session("report_status") = "4" then Response.Write " selected" end if%> value="4" <% if session("report_status") = "0" then Response.Write " disabled" end if%>>Received</option>   
                <option <% if session("report_status") = "2" then Response.Write " selected" end if%> value="2" <% if session("report_status") = "0" then Response.Write " disabled" end if%>>Waiting for parts</option>
                <option <% if session("report_status") = "3" then Response.Write " selected" end if%> value="3" <% if session("report_status") = "0" then Response.Write " disabled" end if%>>To be invoiced</option>
                <option <% if session("report_status") = "0" then Response.Write " selected" else Response.Write " disabled" end if%> value="0" style="color:green">Completed / Exported</option>
              </select></td>
          </tr>
        </table>
        <p>
          <input type="hidden" name="Action" />
          <input type="submit" value="Update Report" <% if session("report_status") = "0" then Response.Write " disabled" end if%> />
        </p>
      </form></td>
  </tr>
</table>
</body>
</html>