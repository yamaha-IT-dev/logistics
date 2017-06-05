<% response.cookies("current_URL_cookie_logistics") = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?" & Request.Querystring %>
<!--#include file="include/connection_it.asp " -->
<!--#include file="class/clsPack.asp " -->
<% strSection = "pack" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Pack Requests</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="include/javascript.js"></script>
<script language="JavaScript" type="text/javascript">
function submitLogistics(theForm) {
    theForm.Action.value = 'Logistics';

    return true;
}

function submitWarehouse(theForm) {
    theForm.Action.value = "Warehouse";

    return true;
}

function submitWarehouseETA(theForm) {
    // Check the warehouse eta for an empty / blank value
    if(theForm.txtWarehouseETA.value) {
        // A value is present, allow the form to post
        return true;
    } else {
        // The field is empty, highlight the blank field and stop the form posting
        theForm.txtWarehouseETA.classList.add('error');

        return false;
    }
}

function validateUpdatePackComments(theForm) {
    var reason = "";
    var blnSubmit = true;

    reason += validateEmptyField(theForm.txtComments,"Comments");
    reason += validateSpecialCharacters(theForm.txtComments,"Comments");

    if (reason != "") {
        alert("Some fields need correction:\n" + reason);

        blnSubmit = false;
        return false;
    }

    if (blnSubmit == true) {
        theForm.Action.value = 'Update';
        theForm.submit();

        return true;
    }
}

function searchPack() {
    var strSearch = document.forms[0].txtSearch.value;
    var strStatus  = document.forms[0].cboStatus.value;
    document.location.href = 'list_pack.asp?type=search&txtSearch=' + strSearch + '&cboStatus=' + strStatus;
}

function resetSearch() {
    document.location.href = 'list_pack.asp?type=reset';
}
</script>
</head>
<body>
<%
sub setSearch
    select case Trim(Request("type"))
        case "reset"
            session("pack_search")          = ""
            session("pack_status")          = ""
            session("pack_initial_page")    = 1
        case "search"
            session("pack_search")          = Trim(Request("txtSearch"))
            session("pack_status")          = Trim(Request("cboStatus"))
            session("pack_initial_page")    = 1
    end select
end sub

sub displayPack
    dim iRecordCount
    iRecordCount = 0
    Dim strSearch
    dim strSQL
    dim strStatus
    dim intRecordCount
    dim strModifiedDate
    dim strTodayDate
    strTodayDate = FormatDateTime(Date())

    if session("pack_status") = "" then
        session("pack_status") = "1"
    end if

    call OpenDataBase()

    set rs = Server.CreateObject("ADODB.recordset")

    rs.CursorLocation = 3   'adUseClient
    rs.CursorType = 3       'adOpenStatic
    rs.PageSize = 100

    strSQL = "SELECT * FROM logistic_pack "
    strSQL = strSQL & " WHERE  (packEmail LIKE '%" & session("pack_search") & "%' "
    strSQL = strSQL & "         OR packName LIKE '%" & session("pack_search") & "%' "
    strSQL = strSQL & "         OR packComments LIKE '%" & session("pack_search") & "%') "
    strSQL = strSQL & "     AND packstatus LIKE '%" & session("pack_status") & "%' "
    strSQL = strSQL & " ORDER BY packDateCreated DESC"

    rs.Open strSQL, conn

    'Response.Write strSQL

    intPageCount = rs.PageCount
    intRecordCount = rs.recordcount

    Select Case Request("Action")
        case "<<"
            intpage = 1
            session("pack_initial_page") = intpage
        case "<"
            intpage = Request("intpage") - 1
            session("pack_initial_page") = intpage

            if session("pack_initial_page") < 1 then session("pack_initial_page") = 1
        case ">"
            intpage = Request("intpage") + 1
            session("pack_initial_page") = intpage

            if session("pack_initial_page") > intPageCount then session("pack_initial_page") = IntPageCount
        Case ">>"
            intpage = intPageCount
            session("pack_initial_page") = intpage
    end select

    strDisplayList = ""

    if not DB_RecSetIsEmpty(rs) Then

        rs.AbsolutePage = session("pack_initial_page")

        For intRecord = 1 To rs.PageSize
            if (DateDiff("d",rs("packDateModified"), strTodayDate) = 0) OR (DateDiff("d",rs("packDateCreated"), strTodayDate) = 0) then
                if iRecordCount Mod 2 = 0 then
                    strDisplayList = strDisplayList & "<tr class=""updated_today"">"
                else
                    strDisplayList = strDisplayList & "<tr class=""updated_today_2"">"
                end if
            else
                if iRecordCount Mod 2 = 0 then
                    strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
                else
                    strDisplayList = strDisplayList & "<tr class=""innerdoc_2"">"
                end if
            end if

            'strDisplayList = strDisplayList & "<td><a href=""update_pack.asp?id=" & rs("packID") & """><img src=""images/icon_view.png"" border=""0""></a></td>"
            strDisplayList = strDisplayList & "<td nowrap>" & Lcase(rs("packEmail"))
            if DateDiff("d",rs("packDateCreated"), strTodayDate) = 0 then
                strDisplayList = strDisplayList & " <img src=""images/icon_new.gif"" border=""0"">"
            end if

            if rs("packPriority") = 1 then
                strDisplayList = strDisplayList & " <img src=""images/icon_priority.gif"" border=""0"">"
            end if

            strDisplayList = strDisplayList & "</td>"
            strDisplayList = strDisplayList & "<td>" & WeekDayName(WeekDay(rs("packDateCreated"))) & ", " & FormatDateTime(rs("packDateCreated"),1) & "</td>"
            strDisplayList = strDisplayList & "<td>" & rs("packName") & " - " & rs("packQty") & "</td>"
            strDisplayList = strDisplayList & "<td>" & rs("packName2") & " - " & rs("packQty2") & "</td>"
            strDisplayList = strDisplayList & "<td>" & rs("packName3") & " - " & rs("packQty3") & "</td>"
            strDisplayList = strDisplayList & "<td>" & rs("packName4") & " - " & rs("packQty4") & "</td>"
            strDisplayList = strDisplayList & "<td>" & rs("packName5") & " - " & rs("packQty5") & "</td>"
            strDisplayList = strDisplayList & "<td>"
            Select Case rs("packLogistics")
                case 1
                    strDisplayList = strDisplayList & "<img src=""images/tick.gif"" title=""" & rs("packLogisticsDate") &  """>"
                case 0
                    strDisplayList = strDisplayList & "<form method=""post"" name=""form_approve"" id=""form_approve"" onsubmit=""return submitLogistics(this)"">"
                    strDisplayList = strDisplayList & "<input type=""hidden"" name=""Action"" value=""Logistics"">"
                    strDisplayList = strDisplayList & "<input type=""hidden"" name=""packID"" value=""" & rs("packID") & """>"
                    strDisplayList = strDisplayList & "<input type=""hidden"" name=""packName"" value=""" & rs("packName") & """>"
                    strDisplayList = strDisplayList & "<input type=""hidden"" name=""packQty"" value=""" & rs("packQty") & """>"
                    strDisplayList = strDisplayList & "<input type=""hidden"" name=""packName2"" value=""" & rs("packName2") & """>"
                    strDisplayList = strDisplayList & "<input type=""hidden"" name=""packQty2"" value=""" & rs("packQty2") & """>"
                    strDisplayList = strDisplayList & "<input type=""hidden"" name=""packName3"" value=""" & rs("packName3") & """>"
                    strDisplayList = strDisplayList & "<input type=""hidden"" name=""packQty3"" value=""" & rs("packQty3") & """>"
                    strDisplayList = strDisplayList & "<input type=""hidden"" name=""packName4"" value=""" & rs("packName4") & """>"
                    strDisplayList = strDisplayList & "<input type=""hidden"" name=""packQty4"" value=""" & rs("packQty4") & """>"
                    strDisplayList = strDisplayList & "<input type=""hidden"" name=""packName5"" value=""" & rs("packName5") & """>"
                    strDisplayList = strDisplayList & "<input type=""hidden"" name=""packQty5"" value=""" & rs("packQty5") & """>"
                    strDisplayList = strDisplayList & "<input type=""hidden"" name=""packPriority"" value=""" & rs("packPriority") & """>"
                    strDisplayList = strDisplayList & "<input type=""submit"" value=""Confirm"" style=""color:green"" />"
                    strDisplayList = strDisplayList & "</form>"
            end select
            strDisplayList = strDisplayList & "</td>"
            strDisplayList = strDisplayList & "<td>"
            Select Case rs("packWarehouse")
                case 1
                    strDisplayList = strDisplayList & "<img src=""images/tick.gif"" title=""" & rs("packWarehouseDate") &  """>"
                case 0
                    strDisplayList = strDisplayList & "<form method=""post"" name=""form_approve"" id=""form_approve"" onsubmit=""return submitWarehouse(this)"">"
                    strDisplayList = strDisplayList & "<input type=""hidden"" name=""Action"" value=""Warehouse"">"
                    strDisplayList = strDisplayList & "<input type=""hidden"" name=""packID"" value=""" & rs("packID") & """>"
                    strDisplayList = strDisplayList & "<input type=""hidden"" name=""packName"" value=""" & rs("packName") & """>"
                    strDisplayList = strDisplayList & "<input type=""hidden"" name=""packQty"" value=""" & rs("packQty") & """>"
                    strDisplayList = strDisplayList & "<input type=""hidden"" name=""packName2"" value=""" & rs("packName2") & """>"
                    strDisplayList = strDisplayList & "<input type=""hidden"" name=""packQty2"" value=""" & rs("packQty2") & """>"
                    strDisplayList = strDisplayList & "<input type=""hidden"" name=""packName3"" value=""" & rs("packName3") & """>"
                    strDisplayList = strDisplayList & "<input type=""hidden"" name=""packQty3"" value=""" & rs("packQty3") & """>"
                    strDisplayList = strDisplayList & "<input type=""hidden"" name=""packName4"" value=""" & rs("packName4") & """>"
                    strDisplayList = strDisplayList & "<input type=""hidden"" name=""packQty4"" value=""" & rs("packQty4") & """>"
                    strDisplayList = strDisplayList & "<input type=""hidden"" name=""packName5"" value=""" & rs("packName5") & """>"
                    strDisplayList = strDisplayList & "<input type=""hidden"" name=""packQty5"" value=""" & rs("packQty5") & """>"
                    strDisplayList = strDisplayList & "<input type=""hidden"" name=""packPriority"" value=""" & rs("packPriority") & """>"
                    strDisplayList = strDisplayList & "<input type=""hidden"" name=""packEmail"" value=""" & rs("packEmail") & """>"
                    strDisplayList = strDisplayList & "<input type=""submit"" value=""Confirm"" "
                    if rs("packLogistics") = 0 then
                        strDisplayList = strDisplayList & " disabled "
                    end if
                    strDisplayList = strDisplayList & " style=""color:green"" />"
                    strDisplayList = strDisplayList & "</form>"
            end select
            strDisplayList = strDisplayList & "</td>"
            strDisplayList = strDisplayList & "<td>"
            If IsNull(rs("packWarehouseETA")) Or IsEmpty(rs("packWarehouseETA")) Then
                strDisplayList = strDisplayList & "<form method=""post"" name=""form_approve"" id=""form_approve"" onsubmit=""return submitWarehouseETA(this)"">"
                strDisplayList = strDisplayList & "<input type=""hidden"" name=""Action"" value=""WarehouseETA"">"
                strDisplayList = strDisplayList & "<input type=""hidden"" name=""packID"" value=""" & rs("packID") & """>"
                strDisplayList = strDisplayList & "<input type=""text"" id=""txtWarehouseETA"" name=""txtWarehouseETA"" maxlength=""10"" size=""10"">"
                strDisplayList = strDisplayList & "<input type=""submit"" value=""Add"" style=""color:green;"">"
                strDisplayList = strDisplayList & "</form>"
            Else
                strDisplayList = strDisplayList & WeekDayName(WeekDay(rs("packWarehouseETA"))) & ", " & FormatDateTime(rs("packWarehouseETA"),1)
            End If
            strDisplayList = strDisplayList & "</td>"
            select case rs("packStatus")
                case 1
                    strDisplayList = strDisplayList & "<td>Open"
                case 2
                    strDisplayList = strDisplayList & "<td align=""orange_text"">Cancelled"
                case else
                    strDisplayList = strDisplayList & "<td class=""green_text"">Completed"
            end select
            strDisplayList = strDisplayList & "</td>"
            strDisplayList = strDisplayList & "<form method=""post"" name=""form_update"" id=""form_update"" onsubmit=""return validateUpdatePackComments(this)"">"
            strDisplayList = strDisplayList & "<input type=""hidden"" name=""Action"" value=""Update"">"
            strDisplayList = strDisplayList & "<input type=""hidden"" name=""packID"" value=""" & trim(rs("packID")) & """>"
            strDisplayList = strDisplayList & "<td><input type=""text"" id=""txtComments"" name=""txtComments"" maxlength=""30"" size=""30"" value=""" & rs("packComments") & """ ></td>"
            strDisplayList = strDisplayList & "<td><input type=""submit"" value=""Update"" /></td>"
            strDisplayList = strDisplayList & "</form>"
            strDisplayList = strDisplayList & "<td nowrap>" & rs("packModifiedBy") & ""

            if IsNull(rs("packModifiedBy")) then
                strDisplayList = strDisplayList & "NA</td>"
            else
                strDisplayList = strDisplayList & " - " & WeekDayName(WeekDay(rs("packDateModified"))) & ", " & FormatDateTime(rs("packDateModified"),1) & "</td>"
            end if

            strDisplayList = strDisplayList & "</td>"
            strDisplayList = strDisplayList & "<td><a onclick=""return confirm('Are you sure you want to delete " & rs("packName") & " ?');"" href='delete_pack.asp?id=" & rs("packID") & "'><img src=""images/btn_delete.gif"" border=""0""></a></td>"
            strDisplayList = strDisplayList & "</tr>"

            rs.movenext
            iRecordCount = iRecordCount + 1
            If rs.EOF Then Exit For
        next

    else
        strDisplayList = "<tr class=""innerdoc""><td colspan=""14"" align=""center"">No Packs found.</td></tr>"
    end if

    strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
    strDisplayList = strDisplayList & "<td colspan=""15"" class=""recordspaging"">"
    strDisplayList = strDisplayList & "<form name=""MovePage"" action=""list_pack.asp"" method=""post"">"
    strDisplayList = strDisplayList & "<input type=""hidden"" name=""intpage"" value=" & session("pack_initial_page") & ">"

    if session("pack_initial_page") = 1 then
        strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value=""<<"">"
        strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value=""<"">"
    else
        strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value=""<<"">"
        strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value=""<"">"
    end if
    if session("pack_initial_page") = intpagecount or intRecordCount = 0 then
        strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value="">"">"
        strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value="">>"">"
    else
        strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value="">"">"
        strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value="">>"">"
    end if
    strDisplayList = strDisplayList & "<input type=""hidden"" name=""txtSearch"" value=" & strSearch & ">"
    strDisplayList = strDisplayList & "<input type=""hidden"" name=""cboState"" value=" & strState & ">"
    strDisplayList = strDisplayList & "<br />"
    strDisplayList = strDisplayList & "Page: " & session("pack_initial_page") & " to " & intpagecount
    strDisplayList = strDisplayList & "<br />"
    strDisplayList = strDisplayList & "<h2>Search results: " & intRecordCount & " packs.</h2>"
    strDisplayList = strDisplayList & "</form>"
    strDisplayList = strDisplayList & "</td>"
    strDisplayList = strDisplayList & "</tr>"

    call CloseDataBase()
end sub

sub main
    call UTL_validateLogin
    call setSearch

    if trim(session("pack_initial_page")) = "" then
        session("pack_initial_page") = 1
    end if

    if Request.ServerVariables("REQUEST_METHOD") = "POST" then
        packID              = Trim(Request("packID"))
        packName            = Trim(Request("packName"))
        packQty             = Trim(Request("packQty"))
        packName2           = Trim(Request("packName2"))
        packQty2            = Trim(Request("packQty2"))
        packName3           = Trim(Request("packName3"))
        packQty3            = Trim(Request("packQty3"))
        packName4           = Trim(Request("packName4"))
        packQty4            = Trim(Request("packQty4"))
        packName5           = Trim(Request("packName5"))
        packQty5            = Trim(Request("packQty5"))
        packPriority        = Trim(Request("packPriority"))
        packEmail           = Trim(Request("packEmail"))
        packComments        = Replace(Request.Form("txtComments"),"'","''")
        packWarehouseETA    = Trim(Request.Form("txtWarehouseETA"))

        select case Trim(Request("Action"))
            case "Logistics"
                call logisticsConfirmPack(packID, packName, packQty, packName2,packQty2,packName3,packQty3,packName4,packQty4,packName5,packQty5, packPriority, session("UsrUserName"))
            case "Warehouse"
                call warehouseConfirmPack(packID, packName, packQty, packName2,packQty2,packName3,packQty3,packName4,packQty4,packName5,packQty5, packPriority, session("UsrUserName"), packEmail)
            case "Update"
                call updatePackComments(packID, packComments, session("UsrUserName"))
            case "WarehouseETA"
                call warehouseSetETA(packID, packWarehouseETA, session("UsrUserName"))
        end select
    end if

    call displayPack
end sub

call main

Dim rs, intPageCount, intpage, intRecord, strDisplayList
Dim packID, packName, packQty, packName2,packQty2,packName3,packQty3,packName4,packQty4,packName5,packQty5, packPriority, packEmail, packComments
%>
<table cellspacing="0" cellpadding="0" align="center" class="full_size_table" border="0">
    <!-- #include file="include/header.asp" -->
    <tr>
        <td class="first_content">
            <table cellpadding="5" cellspacing="0" border="0">
                <tr>
                    <td valign="top">
                        <img src="images/icons/icon_pack.jpg" border="0" alt="Cancelled Order" />
                    </td>
                    <td valign="top">
                        <p><img src="images/icon_excel.jpg" border="0" /> <a href="export_pack.asp">Export</a></p>
                    </td>
                    <td valign="top">
                        <div class="alert alert-search">
                        <form name="frmSearch" id="frmSearch" action="list_pack.asp?type=search" method="post" onsubmit="searchPack()">
                            <h3>Search Parameters:</h3>
                            Pack Name / Created by:
                            <input type="text" name="txtSearch" size="25" value="<%= request("txtSearch") %>" maxlength="20" />
                            <select name="cboStatus" onchange="searchPack()">
                                <option <% if session("pack_status") = "1" then Response.Write " selected" end if%> value="1">Open</option>
                                <option <% if session("pack_status") = "0" then Response.Write " selected" end if%> value="0">Completed</option>
                            </select>
                            <input type="button" name="btnSearch" value="Search" onclick="searchPack()" />
                            <input type="button" name="btnReset" value="Reset" onclick="resetSearch()" />
                        </form>
                        </div>
                    </td>
                </tr>
            </table>
            <table cellspacing="0" cellpadding="8" class="database_records">
                <thead>
                    <tr>
                        <td>Created by</td>
                        <td>Date created</td>
                        <td>Pack 1</td>
                        <td>Pack 2</td>
                        <td>Pack 3</td>
                        <td>Pack 4</td>
                        <td>Pack 5</td>
                        <td>1. Logistics</td>
                        <td>2. Warehouse</td>
                        <td>3. Warehouse ETA</td>
                        <td>Status</td>
                        <td>Comments</td>
                        <td>&nbsp;</td>
                        <td>Last modified</td>
                        <td>&nbsp;</td>
                    </tr>
                </thead>
                <tbody>
                    <%= strDisplayList %>
                </tbody>
            </table>
        </td>
    </tr>
</table>

<script type="text/javascript" src="include/jquery-1.12.4.js"></script>
<script type="text/javascript" src="include/moment.js"></script>
<script type="text/javascript" src="include/pikaday.js"></script>
<script type="text/javascript" src="include/pikaday.jquery.js"></script>
<script type="text/javascript">
    $(function() {
        // Attach the datepicker to all the Warehouse ETA inputs
        $("input[name*='txtWarehouseETA']").pikaday({
            firstDay: 1,
            format: 'DD/MM/YYYY',
            minDate: new Date(2000, 0, 1),
            maxDate: new Date(2020, 12, 31),
            yearRange: [2000,2020]
        });
    });
</script>
</body>
</html>