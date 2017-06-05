<% response.cookies("current_URL_cookie_logistics") = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?" & Request.Querystring %>
<!--#include file="include/connection_it.asp " -->
<!--#include file="class/clsGoodsReturn.asp " -->
<!--#include file="class/clsPallet.asp " -->
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

    reason += validateEmptyField(theForm.txtRepairReport,"Repair Report");
    reason += validateSpecialCharacters(theForm.txtRepairReport,"Repair Report");
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

    rs.CursorLocation = 3   'adUseClient
    rs.CursorType = 3       'adOpenStatic

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
' To check whether GRA Report already exists or not
'----------------------------------------------------------------------------------------
sub checkGraReportExist
    dim strSQL
    session("id") = request("id")

    'dim blnReportExists

    set rs = Server.CreateObject("ADODB.recordset")

    rs.CursorLocation = 3   'adUseClient
    rs.CursorType = 3       'adOpenStatic

    Call OpenDataBase()

    strSQL = "SELECT * FROM yma_gra_report WHERE gra_no = '" & session("id") & "' and line_no = '" & session("line_no") & "' and item = '" & session("item_name") & "'"

    'Response.Write strSQL & "<br>"

    rs.Open strSQL, conn

    if not DB_RecSetIsEmpty(rs) Then
        'Yes, report exists
        session("gra_report_exists") = 1
    else
        session("gra_report_exists") = 0
    end if
    'Response.Write session("gra_report_exists")

    call CloseDataBase()
end sub

'----------------------------------------------------------------------------------------
' Add GRA Report to Database
'----------------------------------------------------------------------------------------

sub addGraReport
    dim strSQL

    dim intID
    'dim intLine
    'dim strItem
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

    intID               = Trim(Request("id"))
    'intLine            = Trim(Request("line"))
    'strItem            = Trim(Request("item"))
    strRepairReport     = Replace(Request.Form("txtRepairReport"),"'","''")
    intLabour           = Trim(Request.Form("txtLabour"))
    intParts            = Trim(Request.Form("txtParts"))
    intGST              = Trim(Request.Form("txtGST"))
    intTotalCost        = Trim(Request.Form("txtTotalCost"))
    'strDateReceived    = Trim(Request.Form("txtDateReceived"))
    strDateRepaired     = Trim(Request.Form("txtDateRepaired"))
    strDestination      = Trim(Request.Form("cboDestination"))
    strPalletNo         = Trim(Request.Form("cboPalletNo"))
    intInvoiceExported  = Trim(Request.Form("cboInvoiceExported"))
    strComments         = Replace(Request.Form("txtComments"),"'","''")
    intStatus           = Trim(Request.Form("cboStatus"))

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
    strSQL = strSQL & "'" & intID & "',"
    strSQL = strSQL & "'" & session("line_no") & "',"
    strSQL = strSQL & "'" & session("item_name") & "',"
    strSQL = strSQL & "'" & session("serial_no") & "',"
    strSQL = strSQL & "'" & session("gra_dealer_code") & "" & session("gra_ship_to_dealer") & "',"
    strSQL = strSQL & "'" & session("gra_warranty_code") & "',"
    strSQL = strSQL & "'" & Server.HTMLEncode(strRepairReport) & "',"
    strSQL = strSQL & "CONVERT(money," & intLabour & "),"
    strSQL = strSQL & "CONVERT(money," & intParts & "),"
    strSQL = strSQL & "CONVERT(money," & intGST & "),"
    strSQL = strSQL & "CONVERT(money," & intTotalCost & "),"
    strSQL = strSQL & "CONVERT(datetime,'" & Session("date_received") & "',103),"
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
        Response.Redirect("view_gra.asp?id=" & session("gra_no") & "")
    end if

    call CloseDataBase()
end sub

sub main
    call UTL_validateLogin

    dim intID
    intID = request("id")

    call getGraFromBASE(intID)
    call getGraStatus(intID)
    call getPalletList
    call getWarrantyCode

    Session("line_no")          = Trim(Request("line"))
    'Session("item_name")       = Server.UrlEncode(Request("item"))
    Session("item_name")        = Trim(Request("item"))
    Session("serial_no")        = Trim(Request("serial"))
    Session("date_received")    = Trim(Request("received"))

    if Session("date_received") = "0/0/0" then
        Session("date_received") = "1/1/1900"
    end if

    call checkGraReportExist

    if Request.ServerVariables("REQUEST_METHOD") = "POST" then
        Select Case Trim(Request("Action"))
            case "Add"
                call addGraReport
        end select
    end if
    'response.write "record: " & session("item_record_count")
end sub

dim strMessageText

call main

dim rs
dim intPageCount, intpage, intRecord
dim strDisplayList
dim strPalletList
dim strWarrantyCode
%>
</head>
<body>
<table width="100%" cellpadding="0" cellspacing="0">
  <!-- #include file="include/header.asp" -->
  <tr>
    <td class="first_content"><table cellpadding="5" cellspacing="0" border="0">
        <tr>
          <td valign="top"><a href="list_gra.asp"><img src="images/icon_gra.jpg" border="0" alt="GRA" /></a></td>
          <td valign="top" width="300"><img src="images/backward_arrow.gif" border="0" /> <a href="list_gra.asp">Back to List</a> <img src="images/backward_arrow.gif" border="0" /> <a href="view_gra.asp?id=<%= session("gra_no") %>">View GRA</a>
            <h2>Add GRA Report</h2>
            <font color="green"><%= strMessageText %></font></td>
        </tr>
      </table>
      <table border="0" width="1000">
        <tr>
          <td width="40%" valign="top"><table width="100%" class="white_bordered_table" cellpadding="5" cellspacing="0">
              <tr>
                <td colspan="2" class="item_maintenance_header">GRA no: <u><%= session("gra_no") %></u></td>
              </tr>
              <tr>
                <td width="30%"><strong>Operator:</strong></td>
                <td width="70%"><%= session("gra_operator_name") %></td>
              </tr>
              <tr>
                <td class="column_divider" align="right"><strong>Dealer code:</strong></td>
                <td class="column_divider"><%= session("gra_dealer_code") %><%= session("gra_ship_to_dealer") %></td>
              </tr>
              <tr>
                <td align="right"><strong>Name:</strong></td>
                <td><%= session("gra_dealer_name") %></td>
              </tr>
              <tr>
                <td align="right"><strong>Address:</strong></td>
                <td><%= session("gra_dealer_address") %></td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td><%= session("gra_dealer_city") %></td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td><%= session("gra_dealer_state") %>&nbsp;<%= session("gra_dealer_postcode") %></td>
              </tr>
              <tr>
                <td align="right"><strong>Phone:</strong></td>
                <td><%= session("gra_dealer_phone") %></td>
              </tr>
              <tr>
                <td align="right"><strong>Contact:</strong></td>
                <td><%= session("gra_contact_person") %></td>
              </tr>
              <tr>
                <td class="column_divider"><strong>Plan return date:</strong></td>
                <td class="column_divider"><%= session("gra_day_entered") %>
                  <%  	Select Case session("gra_month_entered") 
							case "1"
								Response.Write(" January ")
							case "2"
								Response.Write(" February ")
							case "3"
								Response.Write(" March ")
							case "4"
								Response.Write(" April ")
							case "5"
								Response.Write(" May ")
							case "6"
								Response.Write(" June ")
							case "7"
								Response.Write(" July ")
							case "8"
								Response.Write(" August ")
							case "9"
								Response.Write(" September ")
							case "10"
								Response.Write(" October ")
							case "11"
								Response.Write(" November ")
							case "12"
								Response.Write(" December ")
						end select %>
                  <%= session("gra_year_entered") %></td>
              </tr>
              <tr>
                <td><strong>Return status:</strong></td>
                <td><%	Select Case session("gra_return_status") 
							case "0"
								Response.Write("Not received ")
							case "1"
								Response.Write("Received, not credited ")
							case "2"
								Response.Write("Credited ")
							case else
								Response.Write session("return_status")
						end select %></td>
              </tr>
              <tr>
                <td><strong>Comments:</strong></td>
                <td><%= session("gra_ext_comment") %> <%= session("gra_int_comment") %></td>
              </tr>
              <tr>
                <td><strong>Warehouse:</strong></td>
                <td><%= session("gra_warehouse") %></td>
              </tr>
              <tr>
                <td><strong>Carrier:</strong></td>
                <td><% 	Select Case session("gra_carrier_code")
							case "J"
								Response.Write("<img src=""images/cope.gif"" border=""0"">")
							case "C"
								Response.Write("Custom pickup")
							case else
								Response.Write session("carrier_code")
						end select %></td>
              </tr>
              <!--<tr>
                <td><strong>Con-note Label Status: </strong></td>
                <td><% 'response.write session("gra_status_label")%>
                </td>
              </tr>-->
            </table></td>
          <td width="60%" valign="top"><% if session("gra_report_exists") = 0 then %>
            <form action="" method="post" name="form_gra" id="form_gra" onsubmit="return validateFormOnSubmit(this)">
              <table width="100%" class="white_bordered_table" cellpadding="5" cellspacing="0">
                <tr>
                  <td colspan="2" class="item_maintenance_header">Report Details - Line no: <u><%= session("line_no") %></u></td>
                </tr>
                <tr>
                  <td valign="top" align="right"><strong>Item:</strong></td>
                  <td><%= session("item_name") %></td>
                </tr>
                <tr>
                  <td valign="top" align="right"><strong>Serial no:</strong></td>
                  <td><%= session("serial_no") %></td>
                </tr>
                <tr>
                  <td valign="top" align="right"><strong>Warranty code:</strong></td>
                  <td><%= session("gra_warranty_code") %></td>
                </tr>
                <tr>
                  <td width="20%" valign="top" class="column_divider">Repair report<span class="mandatory">*</span>:</td>
                  <td width="80%" class="column_divider"><input type="text" id="txtRepairReport" name="txtRepairReport" maxlength="50" size="60" /></td>
                </tr>
                <tr>
                  <td>Labour<span class="mandatory">*</span>:</td>
                  <td>$
                    <input type="text" id="txtLabour" name="txtLabour" maxlength="6" size="8" onfocus="startCalc();" onblur="stopCalc();" /></td>
                </tr>
                <tr>
                  <td>Parts<span class="mandatory">*</span>:</td>
                  <td>$
                    <input type="text" id="txtParts" name="txtParts" maxlength="6" size="8" onfocus="startCalc();" onblur="stopCalc();" /></td>
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
                  <td><%= session("date_received") %> <!--<input type="text" id="txtDateReceived" name="txtDateReceived" maxlength="10" size="10" />
                    <em>DD/MM/YYYY</em>--></td>
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
                  <td>
                  <select name="cboPalletNo">
                  	<%= strPalletList %>
                  </select>
                  </td>
                </tr>
                <tr>
                  <td>Comments:</td>
                  <td>
                  <textarea name="txtComments" id="txtComments" cols="50" rows="5" onKeyDown="limitText(this.form.txtComments,this.form.countdown,200);" 
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
            </form>
            <% else %>
            <h3>Report already exists, please <a href="view_gra.asp?id=<%= session("gra_no") %>">click here</a> to go back.</h3>
<% end if %></td>
        </tr>
      </table></td>
  </tr>
</table>
<script type="text/javascript" src="include/moment.js"></script>
<script type="text/javascript" src="include/pikaday.js"></script>
<script type="text/javascript">
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