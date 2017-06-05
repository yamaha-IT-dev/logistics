<% response.cookies("current_URL_cookie_logistics") = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?" & Request.Querystring %>
<!--#include file="include/connection_it.asp " -->
<!--#include file="class/clsComment.asp " -->
<!--#include file="class/clsDeduction.asp " -->
<!--#include file="class/clsGoodsReturn.asp " -->
<!--#include file="class/clsGoodsReturnLineNo.asp " -->
<!--#include file="class/clsGoodsReturnCheck.asp " -->
<% strSection = "gra" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>View GRA</title>
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="include/generic_form_validations.js"></script>
<script type="text/javascript" src="include/javascript.js"></script>
<script language="JavaScript" type="text/javascript">
function submitComment(theForm) {
    var reason = "";
    var blnSubmit = true;

    reason += validateEmptyField(theForm.txtComment,"Comment");

    if (reason != "") {
        alert("Some fields need correction:\n" + reason);

        blnSubmit = false;
        return false;
    }

    if (blnSubmit == true) {
        theForm.Action.value = 'Comment';

        return true;
    }
}

function submitDeduction(theForm) {
    var reason = "";
    var blnSubmit = true;

    reason += validateEmptyField(theForm.cboDeduction,"Deduction");
    //reason += validateEmptyField(theForm.txtDeductionName,"Deduction");
    //reason += validateNumeric(theForm.txtDeductionAmount,"Amount");
    //reason += validateSpecialCharacters(theForm.txtDeductionAmount,"Amount");

    if (reason != "") {
        alert("Some fields need correction:\n" + reason);

        blnSubmit = false;
        return false;
    }

    if (blnSubmit == true) {
        theForm.Action.value = 'Deduction';

        return true;
    }
}

function submitConnote(theForm) {
    var reason = "";
    var blnSubmit = true;

    reason += validateEmptyField(theForm.txtGraConnote,"GRA Con-note");
    reason += validateSpecialCharacters(theForm.txtGraConnote,"GRA Con-note");

    if (reason != "") {
        alert("Some fields need correction:\n" + reason);

        blnSubmit = false;
        return false;
    }

    if (blnSubmit == true) {
        theForm.Action.value = 'Connote';

        return true;
    }
}

function submitUpdateConnote(theForm) {
    var reason = "";
    var blnSubmit = true;

    reason += validateEmptyField(theForm.txtGraConnote,"GRA Con-note");
    reason += validateSpecialCharacters(theForm.txtGraConnote,"GRA Con-note");

    if (reason != "") {
        alert("Some fields need correction:\n" + reason);

        blnSubmit = false;
        return false;
    }

    if (blnSubmit == true) {
        theForm.Action.value = 'Update Connote';

        return true;
    }
}
</script>
<%
'----------------------------------------------------------------------------------------
' 2. List GRA Details from BASE
'----------------------------------------------------------------------------------------
sub displayGraItems
    dim iLineCount
    iLineCount = 0
    dim iRecordCount
    iRecordCount = 0
    dim intID
    intID = request("id")

    dim strSQL
    dim intRecordCount
    dim strTodayDate

    strTodayDate = FormatDateTime(Date())

    call OpenBaseDataBase()

    set rs = Server.CreateObject("ADODB.recordset")

    rs.CursorLocation   = 3     'adUseClient
    rs.CursorType       = 3     'adOpenStatic
    rs.PageSize         = 100

    if session("item_status") = "" then
        session("item_status") = "1"
    end if

    strSQL = "SELECT DISTINCT BUHYNO, BUHYGY, BUSOSC, BUSTKS, BUHYSU, BUHUSU, BUSIBN, BURIYC, BUAHJY, BUAHJM, BUAHJD, "
    strSQL = strSQL & " cast(BUAHYY as varchar(4)) || right('00'||cast(BUAHYM as varchar(2)),2) || right('00'||cast(BUAHYD as varchar(2)),2) return_date, "
    strSQL = strSQL & "	BUHSCC, BUOINN, BUOING, BUKNGK, BURTST, BUSKN2, BUCLMN, BUSTJN, BUCLMN, YCKBME, B6INNO "
    strSQL = strSQL & " 	FROM BFUEP "
    strSQL = strSQL & "			LEFT OUTER JOIN YFCMP ON TRIM(BURIYC) = TRIM(YCKBCD) "
    strSQL = strSQL & "			LEFT OUTER JOIN BF6EP ON TRIM(B6OINN) = TRIM(BUOINN) and TRIM(BUOINN) > 0 and  TRIM(B6OING) = TRIM(BUOING) and TRIM(BUSIBN) = TRIM(B6SIBN)"
    'strSQL = strSQL & "		WHERE BUSKKI <> 'D' " 'Removed 16 Oct 2014 due to Archived
    strSQL = strSQL & "		WHERE  "
    strSQL = strSQL & "			 YCSKKI <> 'D' "
    strSQL = strSQL & "			AND BUSKKI <> 'D' "
    'strSQL = strSQL & "			AND (YCSKKI <> 'D' OR (YCSKKI = 'D' AND YCSPEC = 'AC')) "
    'strSQL = strSQL & "			AND (YMPMID like 'LOG%' or YMPMID like 'INT%') "
    strSQL = strSQL & "			AND YCKBID = 'RIYC' "
    strSQL = strSQL & "			AND BUHYNO = " & intID

    rs.Open strSQL, conn

    intPageCount = rs.PageCount
    intRecordCount = rs.recordcount

    strDisplayList = ""

    if not DB_RecSetIsEmpty(rs) Then

        For intRecord = 1 To rs.PageSize 
            if iRecordCount Mod 2 = 0 then
                strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
                session("item_record_count") = 0
            else
                strDisplayList = strDisplayList & "<tr class=""innerdoc_2"">"
                session("item_record_count") = 1
            end if

            if trim(rs("BUSTJN")) = "00" then
                strDisplayList = strDisplayList & "<td></td>"
            else
                iLineCount = iLineCount + 1	
                strDisplayList = strDisplayList & "<td><a href=add_gra_report.asp?id=" & trim(rs("BUHYNO")) & "&line=" & iLineCount & "&item=" & trim(rs("BUSOSC")) & "&serial=" & trim(rs("BUSIBN")) & "&received=" & rs("BUAHJD") & "/" & rs("BUAHJM") & "/" & rs("BUAHJY") & ">Add Report</a></td>"

            end if

            if trim(rs("BUSTJN")) = "00" then
                strDisplayList = strDisplayList & "<td></td>"
            else
                strDisplayList = strDisplayList & "<td>" & iLineCount & "</td>"
            end if

            'Item
            strDisplayList = strDisplayList & "<td>" & trim(rs("BUSOSC")) & " "
            if trim(rs("BUSTJN")) = "00" then
                strDisplayList = strDisplayList & " <img src=""images/icon_set-item.gif"" border=""0"">"
            end if

            if trim(rs("BUSTKS")) = "1" then
                strDisplayList = strDisplayList & " <img src=""images/bullet_component.gif"" border=""0"">"
            end if
            strDisplayList = strDisplayList & "</td>"

            'Expected return qty
            strDisplayList = strDisplayList & "<td>" & trim(rs("BUHYSU")) & "</td>"

            'Qty received
            strDisplayList = strDisplayList & "<td>" & trim(rs("BUHUSU")) & "</td>"

            'Serial no
            strDisplayList = strDisplayList & "<td>" & trim(rs("BUSIBN")) & "</td>"

            'Return code
            strDisplayList = strDisplayList & "<td>" & trim(rs("YCKBME")) & "</td>"

            'Condition
            strDisplayList = strDisplayList & "<td>"
            Select Case trim(rs("BUHSCC"))
                case "G"
                    strDisplayList = strDisplayList & "Intact new"
                case "Q"
                    strDisplayList = strDisplayList & "Quality check"
                case "R"
                    strDisplayList = strDisplayList & "Goods faulty"
                case else
                    strDisplayList = strDisplayList & trim(rs("BTUSGR"))
            end select
            strDisplayList = strDisplayList & "</td>"

            'Original invoice
            strDisplayList = strDisplayList & "<td>" & trim(rs("BUOINN")) & "</td>"

            'Credit note
            strDisplayList = strDisplayList & "<td>" & rs("B6INNO") & "</td>"

            'Amount
            strDisplayList = strDisplayList & "<td>$" & rs("BUKNGK") & "</td>"

            'Claim no
            strDisplayList = strDisplayList & "<td>" & rs("BUCLMN") & "</td>"

            'Received Date
            strDisplayList = strDisplayList & "<td>" & rs("BUAHJD") & "/" & rs("BUAHJM") & "/" & rs("BUAHJY") & "</td>"

            strDisplayList = strDisplayList & "</tr>"

            rs.movenext
            iRecordCount = iRecordCount + 1
            If rs.EOF Then Exit For
        next

    else
        strDisplayList = "<tr class=""innerdoc""><td colspan=""13"" align=""center"">No items found.</td></tr>"
    end if

    strDisplayList = strDisplayList & "<tr>"
    strDisplayList = strDisplayList & "<td colspan=""13"" align=""center"">"
    strDisplayList = strDisplayList & "<small>"
    strDisplayList = strDisplayList & intRecordCount & " item(s) found."
    strDisplayList = strDisplayList & "</small></td>"
    strDisplayList = strDisplayList & "</tr>"

    call CloseBaseDataBase()
end sub

'----------------------------------------------------------------------------------------
'3. List GRA Reports
'----------------------------------------------------------------------------------------
sub displayGraReports
    dim iRecordCount
    iRecordCount = 0

    dim intTotalLabourCount
    intTotalLabourCount = 0

    dim intTotalPartsCount
    intTotalPartsCount = 0

    dim intTotalCount
    intTotalCount = 0

    dim intID
    intID = request("id")

    dim strSQL
    dim intRecordCount
    dim strTodayDate

    strTodayDate = FormatDateTime(Date())

    call OpenDataBase()

    set rs = Server.CreateObject("ADODB.recordset")

    rs.CursorLocation = 3   'adUseClient
    rs.CursorType = 3       'adOpenStatic
    rs.PageSize = 100

    if session("item_status") = "" then
        session("item_status") = "1"
    end if

    strSQL = "SELECT * FROM yma_gra_report "
    strSQL = strSQL & "	WHERE gra_no = '" & intID & "' "
    strSQL = strSQL & " ORDER BY line_no"

    'response.Write strSQL

    rs.Open strSQL, conn

    intPageCount = rs.PageCount
    intRecordCount = rs.recordcount

    strDisplayReportList = ""

    if not DB_RecSetIsEmpty(rs) Then

        For intRecord = 1 To rs.PageSize 
            if iRecordCount Mod 2 = 0 then
                strDisplayReportList = strDisplayReportList & "<tr class=""innerdoc"">"
                session("item_record_count") = 0
            else
                strDisplayReportList = strDisplayReportList & "<tr class=""innerdoc_2"">"
                session("item_record_count") = 1
            end if
            strDisplayReportList = strDisplayReportList & "<td><a href=""update_gra_report.asp?id=" & rs("report_id") & """>Edit Report</a></td>"
            strDisplayReportList = strDisplayReportList & "<td>" & trim(rs("line_no")) & "</td>"
            strDisplayReportList = strDisplayReportList & "<td>" & trim(rs("item")) & "</td>"
            strDisplayReportList = strDisplayReportList & "<td>" & trim(rs("repair_report")) & " "
            if DateDiff("d",rs("date_created"), strTodayDate) = 0 then
                strDisplayReportList = strDisplayReportList & " <img src=""images/icon_new.gif"" border=""0"">"
            end if
            strDisplayReportList = strDisplayReportList & "</td>"

            strDisplayReportList = strDisplayReportList & "<td>$" & trim(rs("labour")) & "</td>"
            strDisplayReportList = strDisplayReportList & "<td>$" & trim(rs("parts")) & "</td>"
            strDisplayReportList = strDisplayReportList & "<td>$" & trim(rs("gst")) & "</td>"
            strDisplayReportList = strDisplayReportList & "<td>$" & trim(rs("total")) & "</td>"

            strDisplayReportList = strDisplayReportList & "<td>" & trim(rs("date_received")) & "</td>"
            strDisplayReportList = strDisplayReportList & "<td>" & trim(rs("destination")) & "</td>"
            strDisplayReportList = strDisplayReportList & "<td>" & trim(rs("pallet_no")) & "</td>"

            strDisplayReportList = strDisplayReportList & "<td>"
            Select Case rs("status")
                case 1
                    strDisplayReportList = strDisplayReportList & "Open"
                case 2
                    strDisplayReportList = strDisplayReportList & "Waiting for parts"
                case 3
                    strDisplayReportList = strDisplayReportList & "To be invoiced"
                case 4
                    strDisplayReportList = strDisplayReportList & "Received"
                case else
                    strDisplayReportList = strDisplayReportList & "<font color=""green"">Completed / Exported</font>"
            end select
            strDisplayReportList = strDisplayReportList & "</td>"

            strDisplayReportList = strDisplayReportList & "</tr>"

            intTotalLabourCount = intTotalLabourCount + trim(rs("labour"))
            intTotalPartsCount = intTotalPartsCount + trim(rs("parts"))
            intTotalCount = intTotalCount + trim(rs("total"))

            rs.movenext

            iRecordCount = iRecordCount + 1

            If rs.EOF Then Exit For
        next
    else
        strDisplayReportList = "<tr class=""innerdoc""><td colspan=""12"" align=""center"">No reports found.</td></tr>"
    end if

    strDisplayReportList = strDisplayReportList & "<tr><td colspan=""12"" align=""center"">Total Labour: $" & intTotalLabourCount & "<br />"
    strDisplayReportList = strDisplayReportList & "Total Parts: $" & intTotalPartsCount & "<br />"
    strDisplayReportList = strDisplayReportList & "<strong>Grand Total: $" & intTotalCount & "</strong></td></tr>"
    strDisplayReportList = strDisplayReportList & "<tr>"
    strDisplayReportList = strDisplayReportList & "<td colspan=""13"" align=""center"">"
    strDisplayReportList = strDisplayReportList & "<small>"
    strDisplayReportList = strDisplayReportList & intRecordCount & " report(s) found."
    strDisplayReportList = strDisplayReportList & "</small></td>"
    strDisplayReportList = strDisplayReportList & "</tr>"

    call CloseDataBase()
end sub

sub main
    call UTL_validateLogin
    session("gra_connote")      = ""
    session("gra_date_created") = ""
    session("gra_AV_check")     = ""
    session("gra_MPDC_check")   = ""

    dim intID
    intID = request("id")

    dim intDeductionID
    intDeductionID = Request.Form("cboDeduction")

    dim intDeductionLine
    intDeductionLine = Request.Form("cboDeductionLine")

    dim strDeductionComments
    strDeductionComments = Trim(Request.Form("txtDeductionComments"))

    dim strGraConnote
    strGraConnote = Request.Form("txtGraConnote")

    select case Trim(Request("ref"))
        case "gra"
            strBackLink = "list_gra.asp"
        case "report"
            strBackLink = "list_gra_report.asp"
        case "writeoffs"
            strBackLink = "list_gra_report_writeoffs.asp"
        case "exported"
            strBackLink = "list_gra_report_exported.asp"
        case else
            strBackLink = "list_gra.asp"
    end select

    if Request.ServerVariables("REQUEST_METHOD") = "POST" then
        Select Case Trim(Request("Action"))
            case "Comment"
                call addComment(intID,graModuleID)
            case "Deduction"
                call addDeduction(intID,intDeductionID,intDeductionLine,strDeductionComments)
            case "Connote"
                call addGraConnote(intID,strGraConnote)
            case "Update Connote"
                call updateGraConnote(intID,strGraConnote)
        end select
    end if

    call getGraFromBASE(intID)
    call getGraStatus(intID)
    call getGraConnote(intID)
    call displayGraItems
    call displayGraReports
    call listComments(intID,graModuleID)
    call listDeductions(intID)
    call checkAV(intID)
    call checkMPD(intID)
    call getAllDeductionTypeList(intID)
    call getGRALineNo(intID)
end sub

dim strMessageText
dim strBackLink

call main

dim rs
dim intPageCount, intpage, intRecord
dim strDisplayList
dim strDisplayReportList
dim strCommentsList
dim strDeductionList
dim strTotalDeductionList
dim strAllDeductionTypeList
dim strGRALineNoList
dim intTotalDeduction
%>
</head>
<body>
<table width="100%" cellpadding="0" cellspacing="0">
  <!-- #include file="include/header.asp" -->
  <tr>
    <td class="first_content"><table cellpadding="5" cellspacing="0" border="0">
        <tr>
          <td valign="top"><a href="list_gra.asp"><img src="images/icon_gra.jpg" border="0" alt="GRA" /></a></td>
          <td valign="top"><img src="images/backward_arrow.gif" border="0" /> <a href="<%= strBackLink %>">Back to List</a>
            <h2>Goods Return Authorisation</h2></td>
        </tr>
      </table>
      <table border="0" width="100%">
        <tr>
          <td width="50%" valign="top">
          <% if session("gra_not_found") <> "TRUE" then %>
          <table width="100%" class="white_bordered_table" cellpadding="5" cellspacing="0">
              <tr>
                <td colspan="2" class="item_maintenance_header">GRA no: <u><%= session("gra_no") %></u></td>
              </tr>
              <tr>
                <td><strong>Yamaha Operator:</strong></td>
                <td><%= session("gra_operator_code") %></td>
              </tr>
              <tr>
                <td colspan="2" align="left" class="column_divider"><strong>Dealer Details:</strong>
                  <div id="contentstart">
                    <table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr>
                        <td width="30%" align="right">&nbsp;</td>
                        <td width="70%"><u><strong><%= session("gra_dealer_name") %></strong></u> (<%= session("gra_dealer_code") %><%= session("gra_ship_to_dealer") %>)</td>
                      </tr>
                      <tr>
                        <td align="right">&nbsp;</td>
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
                    </table>
                  </div></td>
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
                <td width="30%" class="column_divider"><strong>Plan return date:</strong></td>
                <td width="70%" class="column_divider"><%= session("gra_day_entered") %>
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
							case "R"
								Response.Write("<img src=""images/startrack.jpg"">")
							case "S"
								Response.Write("<img src=""images/startrack.jpg"">")
							case else
								Response.Write session("carrier_code")
						end select %></td>
              </tr>
              <tr>
                <td class="column_divider"><strong>Came through AV portal:</strong></td>
                <td class="column_divider"><% Select Case session("gra_AV_check")
							case 1
								Response.Write("<img src=""images/tick.png"" border=""0"">")
							case 0
								Response.Write("<img src=""images/cross.gif"" border=""0"">")
							case else
								Response.Write "-"
						end select %>
				</td>
              </tr>
              <tr>
                <td><strong>Came through MPD portal:</strong></td>
                <td><% Select Case session("gra_MPD_check")
							case 1
								Response.Write("<img src=""images/tick.png"" border=""0"">")
							case 0
								Response.Write("<img src=""images/cross.gif"" border=""0"">")
							case else
								Response.Write "-"
						end select %></td>
              </tr>
              <!--<tr>
                <td><strong>Con-note Label Status: </strong></td>
                <td><% 'response.write session("gra_status_label")%></td>
              </tr>-->
            </table>
            <% else %>
            <h1>Sorry but GRA <%= request("id") %> not found in BASE.</h1>
            <% end if %>
            <p><img src="images/icon_printer.gif" border="0" /> <a href="javascript:PrintThisPage()">Printable version</a></p>
            <table cellpadding="5" cellspacing="0" class="white_bordered_table">
              <tr>
                <td class="item_maintenance_header">CON-NOTE</td>
              </tr>
              <tr>
                <td><% if len(trim(session("gra_connote"))) = 0 then  %>
                  <form action="" name="form_add_connote" id="form_add_connote" method="post" onSubmit="return submitConnote(this)">
                    <input type="text" name="txtGraConnote" id="txtGraConnote" maxlength="10" size="12" />
                    <input type="hidden" name="Action" />
                    <input type="submit" value="Add Connote" />
                  </form>
                  <% else %>
                  <form action="" name="form_update_connote" id="form_update_connote" method="post" onSubmit="return submitUpdateConnote(this)">
                    <input type="text" name="txtGraConnote" id="txtGraConnote" maxlength="10" size="12" value="<%= session("gra_connote") %>" />
                    <input type="hidden" name="Action" />
                    <input type="submit" value="Update Connote" />
                    <p> <strong>Created by:</strong> <%= session("gra_created_by") %> - <%= displayDateFormatted(session("gra_date_created")) %> <br />
                      <strong>Modified by:</strong> <%= session("gra_modified_by") %> - <%= displayDateFormatted(session("gra_date_modified")) %> </p>
                  </form>
                  <% end if %></td>
              </tr>
            </table>
            <br />
            <h3>Item(s):</h3>
            <table cellspacing="0" cellpadding="5" class="database_records">
              <tr class="innerdoctitle">
                <td>&nbsp;</td>
                <td>Line</td>
                <td>Item</td>
                <!--<td>Set comp. qty</td>-->
                <td>Expected return qty</td>
                <td>Qty received</td>
                <td>Serial</td>
                <td>Return code</td>
                <td>Condition</td>
                <td>Ori invoice</td>
                <td>Credit note</td>
                <td>Amount</td>
                <td style="color:#FF0">Claim #</td>
                <td>Received</td>
              </tr>
              <%= strDisplayList %>
            </table></td>
          <td width="50%" valign="top"><h3>Report Summaries</h3>
            <table cellspacing="0" cellpadding="5" class="database_records">
              <tr class="innerdoctitle">
                <td></td>
                <td>Line</td>
                <td>Item</td>
                <td>Repair report</td>
                <td>Labour</td>
                <td>Parts</td>
                <td>GST</td>
                <td>Total</td>
                <td>Received</td>
                <td>Destination</td>
                <td>Pallet</td>
                <!--<td>Exported</td>-->
                <td>Report status</td>
              </tr>
              <%= strDisplayReportList %>
            </table>
            <br />
            <div id="content_start">
            <table cellpadding="5" cellspacing="0" class="deduction_box" width="100%">
              <tr>
                <td class="item_maintenance_header" colspan="7">Deductions</td>
              </tr>
              <tr>
              <td><strong>Line</strong></td>
              <td><strong>Deduction</strong></td>
              <td><strong>$</strong></td>
              <td><strong>Comments</strong></td>
              <td><strong>Added by</strong></td>
              <td><strong>Created</strong></td>
              <td></td>
              </tr>
              <%= strDeductionList %>
              <tr>
                <td colspan="7"><p align="center"><strong>Total: $
                    <% call sumTotalDeductions(request("id")) %>
                    </strong></p>
                  <form action="" name="form_add_deduction" id="form_add_deduction" method="post" onSubmit="return submitDeduction(this)">
                    <select name="cboDeduction" id="cboDeduction">
                      <%= strAllDeductionTypeList %>
                    </select>
                    <select name="cboDeductionLine" id="cboDeductionLine">
                      <%= strGRALineNoList %>
                    </select>
                    <input type="text" name="txtDeductionComments" id="txtDeductionComments" maxlength="100" size="50" placeholder="Comments" />
                    <input type="hidden" name="Action" />
                    <input type="submit" value="Add" />
                  </form></td>
              </tr>
            </table>
            </div>
            <!--<p><a href="javascript:PrintDeductionList()"><img src="images/icon_printer.gif" border="0" /></a></p>-->
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
      </table></td>
  </tr>
</table>
</body>
</html>