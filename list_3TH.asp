<% response.cookies("current_URL_cookie_logistics") = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?" & Request.Querystring %>
<!--#include file="include/connection_it.asp " -->
<% strSection = "3TH" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>3TH</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="include/javascript.js"></script>
<script type="text/javascript" src="include/jquery.js"></script>
<script type="text/javascript" src="include/main.js"></script>
<script language="JavaScript" type="text/javascript">
function searchItem() {
  var strSearch       = document.forms[0].txtSearch.value;
  var strType         = document.forms[0].cboType.value;
  var strDepartment   = document.forms[0].cboDepartment.value;
  var strStatus       = document.forms[0].cboStatus.value;

  document.location.href = 'list_3TH.asp?type=search&txtSearch=' + strSearch + '&cboType=' + strType + '&cboDepartment=' + strDepartment + '&cboStatus=' + strStatus;
}

function resetSearch() {
  document.location.href = 'list_3TH.asp?type=reset';
}
</script>
</head>
<body>
<%
sub setSearch
  select case Trim(Request("type"))
    case "reset"
      session("3TH_return_search")        = ""
      session("3TH_return_type")          = ""
      session("3TH_return_department")    = ""
      session("3TH_return_status")        = ""
      session("3TH_return_initial_page")  = 1
    case "search"
      session("3TH_return_search")        = trim(Request("txtSearch"))
      session("3TH_return_type")          = request("cboType")
      session("3TH_return_department")    = request("cboDepartment")
      session("3TH_return_status")        = Trim(Request("cboStatus"))
      session("3TH_return_initial_page")  = 1
  end select
end sub

sub displayQuarantine
  dim iRecordCount
  iRecordCount = 0
  dim strSortBy
  dim strSortItem
  dim strSQL
  dim strPageResultNumber
  dim strRecordPerPage
  dim intRecordCount
  dim strTodayDate
  dim strDays

  strTodayDate = FormatDateTime(Date())

  call OpenDataBase()

  set rs = Server.CreateObject("ADODB.recordset")

  rs.CursorLocation = 3   'adUseClient
  rs.CursorType = 3       'adOpenStatic
  rs.PageSize = 100

  if session("3TH_return_status") = "" then
    session("3TH_return_status") = "1"
  end if

  strSQL = "SELECT * FROM yma_3th_return "
  strSQL = strSQL & "WHERE department LIKE '%" & session("3TH_return_department") & "%' "

  if session("3TH_return_type") <> "" then
    strSQL = strSQL & "AND return_type = '" & session("3TH_return_type") & "' "
  end if

  strSQL = strSQL & "AND (item_code LIKE '%" & session("3TH_return_search") & "%' "
  strSQL = strSQL & "OR shipment_no LIKE '%" & session("3TH_return_search") & "%' "
  strSQL = strSQL & "OR description LIKE '%" & session("3TH_return_search") & "%' "
  strSQL = strSQL & "OR carrier LIKE '%" & session("3TH_return_search") & "%' "
  strSQL = strSQL & "OR label_no LIKE '%" & session("3TH_return_search") & "%' "
  strSQL = strSQL & "OR original_connote LIKE '%" & session("3TH_return_search") & "%' "
  strSQL = strSQL & "OR gra LIKE '%" & session("3TH_return_search") & "%' "
  strSQL = strSQL & "OR serial_no LIKE '%" & session("3TH_return_search") & "%' ) "
  strSQL = strSQL & "AND status LIKE '%" & session("3TH_return_status") & "%' "
  strSQL = strSQL & "ORDER BY date_created DESC"

  'Response.Write strSQL & "<br>"

  rs.Open strSQL, conn

  intPageCount = rs.PageCount
  intRecordCount = rs.recordcount

  Select Case Request("Action")
    case "<<"
      intpage = 1
      session("3TH_return_initial_page") = intpage
    case "<"
      intpage = Request("intpage") - 1
      session("3TH_return_initial_page") = intpage
      if session("3TH_return_initial_page") < 1 then session("3TH_return_initial_page") = 1
    case ">"
      intpage = Request("intpage") + 1
      session("3TH_return_initial_page") = intpage
      if session("3TH_return_initial_page") > intPageCount then session("3TH_return_initial_page") = IntPageCount
    case ">>"
      intpage = intPageCount
      session("3TH_return_initial_page") = intpage
  end select

  strDisplayList = ""

  if not DB_RecSetIsEmpty(rs) Then
    rs.AbsolutePage = session("3TH_return_initial_page")

    For intRecord = 1 To rs.PageSize
      strDays = DateDiff("d",rs("date_created"), strTodayDate)

      if (DateDiff("d",rs("date_modified"), strTodayDate) = 0) then
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

      strDisplayList = strDisplayList & "<td nowrap><a href=""update_3TH.asp?id=" & rs("return_id") & """>" & rs("return_id") & "</a></td>"
      strDisplayList = strDisplayList & "<td>"
      Select Case rs("return_type")
        case 1
          strDisplayList = strDisplayList & "Lost in Warehouse"
        case 2
          strDisplayList = strDisplayList & "Lost by Carrier"
        case 3
          strDisplayList = strDisplayList & "Packaging Issue"
        case 4
          strDisplayList = strDisplayList & "Warehouse Variance"
        case 5
          strDisplayList = strDisplayList & "Display Stock"
        case else
          strDisplayList = strDisplayList & "-"
      end select
      strDisplayList = strDisplayList & "</td>"
      strDisplayList = strDisplayList & "<td>" & rs("department") & "</td>"
      strDisplayList = strDisplayList & "<td nowrap>" & strDays & " days</td>"
      strDisplayList = strDisplayList & "<td><strong>" & rs("item_code") & "</strong>"
      if DateDiff("d",rs("date_created"), strTodayDate) = 0 then
        strDisplayList = strDisplayList & " <img src=""images/icon_new.gif"" border=""0"">"
      end if
      strDisplayList = strDisplayList & "</td>"
      strDisplayList = strDisplayList & "<td>" & rs("qty") & "</td>"
      strDisplayList = strDisplayList & "<td>" & rs("description") & "</td>"
      strDisplayList = strDisplayList & "<td>" & rs("label_no") & "</td>"
      strDisplayList = strDisplayList & "<td>" & rs("dealer") & "</td>"
      strDisplayList = strDisplayList & "<td>" & rs("shipment_no") & "</td>"
      strDisplayList = strDisplayList & "<td>" & rs("carrier") & "</td>"
      strDisplayList = strDisplayList & "<td>" & rs("original_connote") & "</td>"
      if rs("date_received") = "01/01/1900" or rs("date_received") = "1/1/1900" then
        strDisplayList = strDisplayList & "<td class=""orange_text"">TBA</td>"
      else
        strDisplayList = strDisplayList & "<td>" & FormatDateTime(rs("date_received"),1) & "</td>"
      end if
      strDisplayList = strDisplayList & "<td>" & rs("serial_no") & "</td>"
      strDisplayList = strDisplayList & "<td>"
      Select Case rs("instruction")
        case "1"
          strDisplayList = strDisplayList & "Update GRA"
        case "2"
          strDisplayList = strDisplayList & "Writeoff Approval Required"
        case else
          strDisplayList = strDisplayList & "-"
      end select
      strDisplayList = strDisplayList & "</td>"
      strDisplayList = strDisplayList & "<td>" & rs("gra") & "</td>"
      if rs("status") = 1 then
        strDisplayList = strDisplayList & "<td class=""blue_text"">Open</td>"
      else
        strDisplayList = strDisplayList & "<td class=""green_text"">Completed</td>"
      end if
      strDisplayList = strDisplayList & "<td><a onclick=""return confirm('Are you sure you want to delete " & rs("item_code") & " ?');"" href='delete_3TH.asp?return_id=" & rs("return_id") & "'><img src=""images/btn_delete.gif"" border=""0""></a></td>"
      strDisplayList = strDisplayList & "</tr>"

      rs.movenext
      iRecordCount = iRecordCount + 1
      If rs.EOF Then Exit For
    next
  else
    strDisplayList = "<tr class=""innerdoc""><td colspan=""18"" align=""center"">No records found.</td></tr>"
  end if

  strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
  strDisplayList = strDisplayList & "<td colspan=""18"" class=""recordspaging"">"
  strDisplayList = strDisplayList & "<form name=""MovePage"" action=""list_3TH.asp"" method=""post"">"
  strDisplayList = strDisplayList & "<input type=""hidden"" name=""intpage"" value=" & session("3TH_return_initial_page") & ">"

  if session("3TH_return_initial_page") = 1 then
    strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value=""<<"">"
    strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value=""<"">"
  else
    strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value=""<<"">"
    strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value=""<"">"
  end if
  if session("3TH_return_initial_page") = intpagecount or intRecordCount = 0 then
    strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value="">"">"
    strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value="">>"">"
  else
    strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value="">"">"
    strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value="">>"">"
  end if
  strDisplayList = strDisplayList & "<input type=""hidden"" name=""txtSearch"" value=" & strSearch & ">"
  strDisplayList = strDisplayList & "<input type=""hidden"" name=""cboDepartment"" value=" & strItemDepartment & ">"
  strDisplayList = strDisplayList & "<input type=""hidden"" name=""cboStatus"" value=" & strStatus & ">"
  strDisplayList = strDisplayList & "<br />"
  strDisplayList = strDisplayList & "Page: " & session("3TH_return_initial_page") & " to " & intpagecount
  strDisplayList = strDisplayList & "<br />"
  strDisplayList = strDisplayList & "<h2>Search results: " & intRecordCount & " records.</h2>"
  strDisplayList = strDisplayList & "</form>"
  strDisplayList = strDisplayList & "</td>"
  strDisplayList = strDisplayList & "</tr>"

  call CloseDataBase()
end sub

sub main
  call UTL_validateLogin
  call setSearch

  if trim(session("3TH_return_initial_page")) = "" then
    session("3TH_return_initial_page") = 1
  end if

  call displayQuarantine
end sub

call main

dim rs
dim intPageCount, intpage, intRecord
dim strDisplayList
%>
<table cellspacing="0" cellpadding="0" align="center" class="full_size_table" border="0">
  <!-- #include file="include/header.asp" -->
  <tr>
    <td class="first_content">
      <table cellpadding="5" cellspacing="0" border="0">
        <tr>
          <td valign="top"><img src="images/icon_return.jpg" border="0" alt="Warehouse Return" /></td>
          <td valign="top"><div class="alert alert-success"><img src="images/add_icon.png" border="0" align="bottom" /> <a href="add_3TH.asp">Add 3TH</a></div>
            <% if Session("UsrLoginRole") = 1 then %>
              <p><img src="images/icon_excel.jpg" border="0" /> <a href="export_3TH.asp?search=<%= session("3TH_return_search") %>&type=<%= session("3TH_return_type") %>&dept=<%= session("3TH_return_department") %>&status=<%= session("3TH_return_status") %>">Export</a></p>
            <% end if %>
          </td>
          <td valign="top">
            <div class="alert alert-search">
              <form name="frmSearch" id="frmSearch" action="list_3TH.asp?type=search" method="post" onsubmit="searchItem()">
                <h3>Search Parameters:</h3>
                Item / Shipment / Description / Connotes / GRA / Serial no :
                <input type="text" name="txtSearch" size="25" value="<%= request("txtSearch") %>" maxlength="20" />
                <select name="cboType" onchange="searchItem()">
                  <option value="">All Types</option>
                  <option <% if session("3TH_return_type") = "1" then Response.Write " selected" end if%> value="1">Lost in Warehouse</option>
                  <option <% if session("3TH_return_type") = "2" then Response.Write " selected" end if%> value="2">Lost by Carrier</option>
                  <option <% if session("3TH_return_type") = "3" then Response.Write " selected" end if%> value="3">Packaging Issue</option>
                  <option <% if session("3TH_return_type") = "4" then Response.Write " selected" end if%> value="4">Warehouse Variance</option>
                  <option <% if session("3TH_return_type") = "5" then Response.Write " selected" end if%> value="5">Display Stock</option>
                </select>
                <select name="cboDepartment" onchange="searchItem()">
                  <option value="">All Depts</option>
                  <option <% if session("3TH_return_department") = "AV" then Response.Write " selected" end if%> value="AV">AV</option>
                  <option <% if session("3TH_return_department") = "MPD" then Response.Write " selected" end if%> value="MPD">MPD</option>
                </select>
                <select name="cboStatus" onchange="searchItem()">
                  <option <% if session("3TH_return_status") = "1" then Response.Write " selected" end if%> value="1">Open</option>
                  <option <% if session("3TH_return_status") = "0" then Response.Write " selected" end if%> value="0">Completed</option>
                </select>
                <input type="button" name="btnSearch" value="Search" onclick="searchItem()" />
                <input type="button" name="btnReset" value="Reset" onclick="resetSearch()" />
              </form>
            </div>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td>
      <table cellspacing="0" cellpadding="8" class="database_records">
        <thead>
          <tr>
            <td>ID</td>
            <td>Type</td>
            <td>Dept</td>
            <td>Lapsed</td>
            <td>Model no</td>
            <td>Qty</td>
            <td>Description</td>
            <td>Label #</td>
            <td>Dealer</td>
            <td>Shipment #</td>
            <td>Carrier</td>
            <td>Connote</td>
            <td>Received</td>
            <td>Serial #</td>
            <td>Instruction</td>
            <td>GRA</td>
            <td>Status</td>
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
</body>
</html>