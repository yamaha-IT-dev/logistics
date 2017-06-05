<% response.cookies("current_URL_cookie_logistics") = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?" & Request.Querystring %>
<!--#include file="include/connection_it.asp " -->
<!--#include file="class/clsGoodsReturn.asp " -->
<!--#include file="class/clsPallet.asp " -->
<!--#include file="class/clsWarrantyCode.asp " -->
<% strSection = "gra" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Add GRA Report</title>
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="include/autoSum.js"></script>
<script type="text/javascript" src="include/usableforms.js"></script>
<script type="text/javascript" src="include/generic_form_validations.js"></script>
<script language="JavaScript" type="text/javascript">
function validateFormOnSubmit(theForm) {
	var reason = "";
	var blnSubmit = true;
	
	reason += validateEmptyField(theForm.txtGraNo,"GRA no");
	reason += validateSpecialCharacters(theForm.txtGraNo,"GRA no");
	
	reason += validateEmptyField(theForm.txtItem,"Item");
	reason += validateSpecialCharacters(theForm.txtItem,"Item");
	
	reason += validateNumeric(theForm.txtLine,"Line");
	
	reason += validateEmptyField(theForm.txtSerialNo,"Serial no");
	reason += validateSpecialCharacters(theForm.txtSerialNo,"Serial no");
	
	reason += validateEmptyField(theForm.txtDealerCode,"Dealer code");
	reason += validateSpecialCharacters(theForm.txtDealerCode,"Dealer code");
	
	reason += validateEmptyField(theForm.cboWarrantyCode,"Warranty code");
	
	reason += validateEmptyField(theForm.txtRepairReport,"Repair report");
	reason += validateSpecialCharacters(theForm.txtRepairReport,"Repair report");
	
	reason += validateNumeric(theForm.txtLabour,"Labour");
	reason += validateSpecialCharacters(theForm.txtLabour,"Labour");
	
	reason += validateNumeric(theForm.txtParts,"Parts");
	reason += validateSpecialCharacters(theForm.txtParts,"Parts");
	//reason += validateSpecialCharacters(theForm.txtComments,"Comments");
	
	if (theForm.cboDestination.value == "Destroy") {
		reason += validateEmptyField(theForm.cboPalletNo,"Pallet");
	}
	
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
<%
'----------------------------------------------------------------------------------------
' Get Warranty Code from BASE
'----------------------------------------------------------------------------------------
sub getWarrantyCode
	dim strSQ
	'dim strWarrantyCode
	
	set rs = Server.CreateObject("ADODB.recordset")

	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic

	Call OpenDataBase()
	
	strSQL = "SELECT gra_warranty_code FROM "
	strSQL = strSQL & " OPENQUERY(AS400, 'SELECT Y3SOSC, Y3GREG FROM YF3ML01 WHERE Y3SOSC = ''" & session("item_name") & "''') "
	strSQL = strSQL & " 	LEFT JOIN yma_gra_warranty_code ON LEFT(Y3GREG, 3) = gra_product_group_code"
	
	'Response.Write strSQL & "<br>"
	
	rs.Open strSQL, conn

	if not DB_RecSetIsEmpty(rs) Then
		session("gra_warranty_code") = trim(rs("gra_warranty_code"))
		'strWarrantyCode = trim(rs("gra_warranty_code"))
    end if
	
	call CloseDataBase()
end sub

'----------------------------------------------------------------------------------------
' Add GRA Report to Database
'----------------------------------------------------------------------------------------

sub addGraReport
	dim strSQL
	
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
	dim intInvoiceExported
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
	intInvoiceExported	= Trim(Request.Form("cboInvoiceExported"))
	strComments 		= Replace(Request.Form("txtComments"),"'","''")
	intStatus			= Trim(Request.Form("cboStatus"))
	
	call OpenDataBase()
		
	strSQL = "INSERT INTO yma_gra_report ("
	strSQL = strSQL & " gra_no, "
	strSQL = strSQL & " line_no, "
	strSQL = strSQL & " item, "
	strSQL = strSQL & " serial_no, "
	strSQL = strSQL & " dealer_code, "
	strSQL = strSQL & " gra_warranty_code, "
	strSQL = strSQL & " repair_report, "
	strSQL = strSQL & " labour, "
	strSQL = strSQL & " parts, "
	strSQL = strSQL & " gst, "
	strSQL = strSQL & " total, "
	strSQL = strSQL & " date_received, "
	strSQL = strSQL & " date_repaired, "
	strSQL = strSQL & " destination, "
	strSQL = strSQL & " pallet_no, "
	strSQL = strSQL & " date_created, "
	strSQL = strSQL & " created_by, "
	strSQL = strSQL & " comments, "
	strSQL = strSQL & " status "
	strSQL = strSQL & ") VALUES ( "
	strSQL = strSQL & "'" & Server.HTMLEncode(intGraNo) & "',"
	strSQL = strSQL & "'" & intLine & "',"
	strSQL = strSQL & "'" & Server.HTMLEncode(strItem) & "',"
	strSQL = strSQL & "'" & Server.HTMLEncode(strSerialNo) & "',"
	strSQL = strSQL & "'" & Server.HTMLEncode(strDealerCode) & "',"
	strSQL = strSQL & "'" & strWarrantyCode & "',"
	strSQL = strSQL & "'" & Server.HTMLEncode(strRepairReport) & "',"
	strSQL = strSQL & "CONVERT(money," & intLabour & "),"
	strSQL = strSQL & "CONVERT(money," & intParts & "),"
	strSQL = strSQL & "CONVERT(money," & intGST & "),"
	strSQL = strSQL & "CONVERT(money," & intTotalCost & "),"
	strSQL = strSQL & "CONVERT(datetime,'" & strDateReceived & "',103),"
	strSQL = strSQL & "CONVERT(datetime,'" & strDateRepaired & "',103),"
	strSQL = strSQL & "'" & strDestination & "',"
	strSQL = strSQL & "'" & Server.HTMLEncode(strPalletNo) & "',getdate(),"
	strSQL = strSQL & "'" & session("UsrUserName") & "',"
	strSQL = strSQL & "'" & Server.HTMLEncode(strComments) & "',"
	strSQL = strSQL & "'" & intStatus & "')"
	
	'response.Write strSQL
	
	on error resume next
	conn.Execute strSQL
	
	if err <> 0 then
		strMessageText = err.description
	else
		Response.Redirect("list_gra_report.asp")
	end if
	
	call CloseDataBase()
end sub

sub main
	call UTL_validateLogin
	call getPalletList
	call getWarrantyCodeList
	
	if Request.ServerVariables("REQUEST_METHOD") = "POST" then	
		Select Case Trim(Request("Action"))
			case "Add"			
				call addGraReport
		end select
	end if
end sub

dim strMessageText

call main

dim rs
dim intPageCount, intpage, intRecord
dim strDisplayList
dim strPalletList
dim strWarrantyCode

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
            <h2>Add GRA Report Manually</h2>
            <font color="green"><%= strMessageText %></font></td>
        </tr>
      </table>
      <form action="" method="post" name="form_gra" id="form_gra" onsubmit="return validateFormOnSubmit(this)">
        <table class="white_bordered_table_small" cellpadding="5" cellspacing="0">
          <tr>
            <td colspan="2" class="item_maintenance_header">Report Details</td>
          </tr>
          <tr>
            <td width="20%">GRA no<span class="mandatory">*</span>:</td>
            <td width="80%"><input type="text" id="txtGraNo" name="txtGraNo" maxlength="10" size="15" /><font color="red">Don't enter the <strong>/XX</strong> in the end, it should be in the <strong>Line</strong></font></td>
          </tr>
          <tr>
            <td>Item<span class="mandatory">*</span>:</td>
            <td><input type="text" id="txtItem" name="txtItem" maxlength="40" size="50" /></td>
          </tr>
          <tr>
            <td>Line<span class="mandatory">*</span>:</td>
            <td><input type="text" id="txtLine" name="txtLine" maxlength="3" size="4" /></td>
          </tr>
          <tr>
            <td>Serial no<span class="mandatory">*</span>:</td>
            <td><input type="text" id="txtSerialNo" name="txtSerialNo" maxlength="20" size="30" /></td>
          </tr>
          <tr>
            <td>Dealer code<span class="mandatory">*</span>:</td>
            <td><input type="text" id="txtDealerCode" name="txtDealerCode" maxlength="20" size="30" /></td>
          </tr>
          <tr>
            <td>Warranty code<span class="mandatory">*</span>:</td>
            <td><select name="cboWarrantyCode"><%= strWarrantyCodeList %></select></td>
          </tr>
          <tr>
            <td>Repair report<span class="mandatory">*</span>:</td>
            <td><input type="text" id="txtRepairReport" name="txtRepairReport" maxlength="50" size="60" /></td>
          </tr>
          <tr>
            <td>Labour<span class="mandatory">*</span>:</td>
            <td>$
              <input type="text" id="txtLabour" name="txtLabour" maxlength="5" size="8" onfocus="startCalc();" onblur="stopCalc();" /></td>
          </tr>
          <tr>
            <td>Parts<span class="mandatory">*</span>:</td>
            <td>$
              <input type="text" id="txtParts" name="txtParts" maxlength="5" size="8" onfocus="startCalc();" onblur="stopCalc();" /></td>
          </tr>
          <tr>
            <td>GST:</td>
            <td>$
              <input type="text" id="txtGST" name="txtGST" maxlength="6" size="8" style="background-color:#CCC" />
              <em>Auto-generated</em></td>
          </tr>
          <tr>
            <td>Total:</td>
            <td>$
              <input type="text" id="txtTotalCost" name="txtTotalCost" maxlength="6" size="8" style="background-color:#CCC" />
              <em>Auto-generated</em></td>
          </tr>
          <tr>
            <td>Date received:</td>
            <td><input type="text" id="txtDateReceived" name="txtDateReceived" maxlength="10" size="10" />
              <em>DD/MM/YYYY</em></td>
          </tr>
          <tr>
            <td>Date repaired:</td>
            <td><input type="text" id="txtDateRepaired" name="txtDateRepaired" maxlength="10" size="10" />
              <em>DD/MM/YYYY</em></td>
          </tr>
          <tr>
            <td>Destination:</td>
            <td><select name="cboDestination">
                <option value="" rel="none">...</option>
                <option value="3H" rel="none">3H</option>
                <option value="3T" rel="none">3T</option>
                <option value="3S" rel="none">3S</option>
                <option value="Destroy" rel="destroy">Destroy</option>
              </select></td>
          </tr>
          <tr rel="destroy">
            <td align="right">Write-off Pallet<span class="mandatory">*</span>:</td>
            <td><select name="cboPalletNo">
                <%= strPalletList %>
              </select></td>
          </tr>
          <tr>
            <td>Comments:</td>
            <td><textarea name="txtComments" id="txtComments" cols="50" rows="5" onKeyDown="limitText(this.form.txtComments,this.form.countdown,200);" 
onKeyUp="limitText(this.form.txtComments,this.form.countdown,200);"></textarea></td>
          </tr>
          <tr class="status_row">
            <td>Status:</td>
            <td><select name="cboStatus">
                <option value="1">Open</option>
                <option value="4">Received</option>
                <option value="2">Waiting for parts</option>
              </select></td>
          </tr>
        </table>
        <p>
          <input type="hidden" name="Action" />
          <input type="submit" value="Add Report" />
        </p>
      </form></td>
  </tr>
</table>
<script type="text/javascript" src="include/moment.js"></script>
<script type="text/javascript" src="include/pikaday.js"></script>
<script type="text/javascript">	
	var picker = new Pikaday(
    {
        field: document.getElementById('txtDateReceived'),		
        firstDay: 1,
        minDate: new Date('2013-01-01'),
        maxDate: new Date('2020-12-31'),
        yearRange: [2013,2020],
		format: 'DD/MM/YYYY'
    });
	
	var picker = new Pikaday(
    {
        field: document.getElementById('txtDateRepaired'),		
        firstDay: 1,
        minDate: new Date('2013-01-01'),
        maxDate: new Date('2020-12-31'),
        yearRange: [2013,2020],
		format: 'DD/MM/YYYY'
    });			
</script>
</body>
</html>