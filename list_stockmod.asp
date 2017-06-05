<% response.cookies("current_URL_cookie_logistics") = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?" & Request.Querystring %>
<!--#include file="include/connection_it.asp " -->
<% strSection = "stock_modification" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Stock Modifications</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="include/javascript.js"></script>
<script language="JavaScript" type="text/javascript">
function searchStockmod() {
    var strSearch = document.forms[0].txtSearch.value;
    var strType  = document.forms[0].cboType.value;
    var strSort  = document.forms[0].cboSort.value;
    document.location.href = 'list_stockmod.asp?type=search&txtSearch=' + strSearch + '&cboType=' + strType + '&cboSort=' + strSort;
}

function resetSearch() {
    document.location.href = 'list_stockmod.asp?type=reset';
}
</script>
</head>
<body>
<%
sub setSearch
    select case Trim(Request("type"))
        case "reset"
            session("stockmod_search")          = ""
            session("stockmod_type")            = ""
            session("stockmod_sort")            = ""
            session("stockmod_initial_page")    = 1
        case "search"
            session("stockmod_search")          = Trim(Request("txtSearch"))
            session("stockmod_type")            = Trim(Request("cboType"))
            session("stockmod_sort")            = Trim(Request("cboSort"))
            session("stockmod_initial_page")    = 1
    end select
end sub

sub displayStockmod
    dim iRecordCount
    iRecordCount = 0
    Dim strSortBy
    dim strSortItem
    Dim strSearch
    dim strSQL
    dim strType
    dim strSort
    dim strPageResultNumber
    dim strRecordPerPage
    dim intRecordCount
    dim strModifiedDate

    dim strTodayDate
    strTodayDate = FormatDateTime(Date())

    call OpenDataBase()

    set rs = Server.CreateObject("ADODB.recordset")

    rs.CursorLocation = 3   'adUseClient
    rs.CursorType = 3       'adOpenStatic
    rs.PageSize = 100

    if session("stockmod_sort") = "" then
        session("stockmod_sort") = "model_name"
    end if

    strSQL = "SELECT * FROM yma_stock_mod "
    strSQL = strSQL & "	WHERE model_type LIKE '%" & session("stockmod_type") & "%' "
    strSQL = strSQL & "		AND status = 1 "
    strSQL = strSQL & "		AND (model_name LIKE '%" & session("stockmod_search") & "%' "
    strSQL = strSQL & "			OR part_no_base LIKE '%" & session("stockmod_search") & "%' "
    strSQL = strSQL & "			OR created_by LIKE '%" & session("stockmod_search") & "%' "
    strSQL = strSQL & "			OR vendor_model_no LIKE '%" & session("stockmod_search") & "%') "
    strSQL = strSQL & "	ORDER BY " & session("stockmod_sort")

    rs.Open strSQL, conn

    'Response.Write strSQL

    intPageCount = rs.PageCount
    intRecordCount = rs.recordcount

    Select Case Request("Action")
        case "<<"
            intpage = 1
            session("stockmod_initial_page") = intpage
        case "<"
            intpage = Request("intpage") - 1
            session("stockmod_initial_page") = intpage

            if session("stockmod_initial_page") < 1 then session("stockmod_initial_page") = 1
        case ">"
            intpage = Request("intpage") + 1
            session("stockmod_initial_page") = intpage

            if session("stockmod_initial_page") > intPageCount then session("stockmod_initial_page") = IntPageCount
        Case ">>"
            intpage = intPageCount
            session("stockmod_initial_page") = intpage
    end select

    strDisplayList = ""

    if not DB_RecSetIsEmpty(rs) Then

        rs.AbsolutePage = session("stockmod_initial_page")

        For intRecord = 1 To rs.PageSize
            if (DateDiff("d",rs("date_modified"), strTodayDate) = 0) OR (DateDiff("d",rs("date_created"), strTodayDate) = 0) then
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
            'strDisplayList = strDisplayList & "<td><a href=""update_stockmod.asp?id=" & rs("stock_id") & """><img src=""images/icon_view.png"" border=""0""></a></td>"
            strDisplayList = strDisplayList & "<td><a href=""update_stockmod.asp?id=" & rs("stock_id") & """>" & rs("model_name") & "</a>"
            if DateDiff("d",rs("date_created"), strTodayDate) = 0 then
                strDisplayList = strDisplayList & " <img src=""images/icon_new.gif"" border=""0"">"
            end if
            strDisplayList = strDisplayList & "</td>"
            strDisplayList = strDisplayList & "<td>"
            If IsNull(rs("takt_timing")) Then
                strDisplayList = strDisplayList & "-"
            Else
                strDisplayList = strDisplayList & rs("takt_timing")
            End If
            strDisplayList = strDisplayList & "</td>"
            strDisplayList = strDisplayList & "<td>" & rs("model_type") & "</td>"
            strDisplayList = strDisplayList & "<td>" & rs("part_no_base") & "</td>"
            strDisplayList = strDisplayList & "<td>" & rs("vendor_model_no") & "</td>"
            strDisplayList = strDisplayList & "<td>"
            if rs("document") = "1" then
                Select Case Session("UsrLoginRole")
                    case 3 'MOL
                        strDisplayList = strDisplayList & "<a href=""ftp://203.221.101.249/Logistics/_stockmods/" & rs("model_name") & ".pdf"" target=""_blank"">View</a>"
                    case 14 'TT Admin
                        strDisplayList = strDisplayList & "<a href=""file:\\172.29.64.6\shipment\_stockmods\" & rs("model_name") & ".pdf"" target=""_blank"">View</a>"
                    case 15 'TT Normal Users
                        strDisplayList = strDisplayList & "<a href=""file:\\172.29.64.6\shipment\_stockmods\" & rs("model_name") & ".pdf"" target=""_blank"">View</a>"
                    case 16 'TT Normal Users
                        strDisplayList = strDisplayList & "<a href=""ftp://yamaha_vic%5CTTLogShipment:ttL0gix@203.221.101.249/Logistics/_stockmods/" & rs("model_name") & ".pdf"" target=""_blank"">View</a>"
                    case 17
                        strDisplayList = strDisplayList & "<a href=""ftp://yamaha_vic%5CTTLogShipment:ttL0gix@203.221.101.249/Logistics/_stockmods/" & rs("model_name") & ".pdf"" target=""_blank"">View</a>"	
                    case else
                        strDisplayList = strDisplayList & "<a href=""file://///YAMMAS22/shipment/_stockmods/" & rs("model_name") & ".pdf"" target=""_blank"">View</a>"
                end select
            else
                strDisplayList = strDisplayList & "..."
            end if
            strDisplayList = strDisplayList & "</td>"
            strDisplayList = strDisplayList & "<td>" & rs("created_by") & " "
            if IsNull(rs("date_created")) then
                strDisplayList = strDisplayList & "NA</td>"
            else
                strDisplayList = strDisplayList & " - " & FormatDateTime(rs("date_created"),2) & "</td>"
            end if
            strDisplayList = strDisplayList & "<td>" & rs("modified_by") & ""

            if IsNull(rs("date_modified")) then
                strDisplayList = strDisplayList & "NA"
            else
                strDisplayList = strDisplayList & " - " & FormatDateTime(rs("date_modified"),2) & ""
            end if

            strDisplayList = strDisplayList & "</td>"
            strDisplayList = strDisplayList & "<td><a onclick=""return confirm('Are you sure you want to delete " & rs("model_name") & " ?');"" href='delete_stockmod.asp?id=" & rs("stock_id") & "'><img src=""images/btn_delete.gif"" border=""0""></a></td>"
            strDisplayList = strDisplayList & "</tr>"

            rs.movenext
            iRecordCount = iRecordCount + 1
            If rs.EOF Then Exit For
        next
    else
        strDisplayList = "<tr class=""innerdoc""><td colspan=""8"" align=""center"">No records found.</td></tr>"
    end if

    strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
    strDisplayList = strDisplayList & "<td colspan=""9"" class=""recordspaging"">"
    strDisplayList = strDisplayList & "<form name=""MovePage"" action=""list_stockmod.asp"" method=""post"">"
    strDisplayList = strDisplayList & "<input type=""hidden"" name=""intpage"" value=" & session("stockmod_initial_page") & ">"

    if session("stockmod_initial_page") = 1 then
        strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value=""<<"">"
        strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value=""<"">"
    else
        strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value=""<<"">"
        strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value=""<"">"
    end if
    if session("stockmod_initial_page") = intpagecount or intRecordCount = 0 then
        strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value="">"">"
        strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value="">>"">"
    else
        strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value="">"">"
        strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value="">>"">"
    end if

    strDisplayList = strDisplayList & "<input type=""hidden"" name=""txtSearch"" value=" & strSearch & ">"
    strDisplayList = strDisplayList & "<input type=""hidden"" name=""cboType"" value=" & strType & ">"
    strDisplayList = strDisplayList & "<input type=""hidden"" name=""cboSort"" value=" & strSort & ">"
    strDisplayList = strDisplayList & "<br />"
    strDisplayList = strDisplayList & "Page: " & session("stockmod_initial_page") & " to " & intpagecount
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

    if trim(session("stockmod_initial_page")) = "" then
        session("stockmod_initial_page") = 1
    end if

    call displayStockmod
end sub

call main

dim rs
dim intPageCount, intpage, intRecord
dim strDisplayList
dim strDealerResultList
dim strStateList
dim strSalesManagerList
%>
<table cellspacing="0" cellpadding="0" align="center" class="full_size_table" border="0">
  <!-- #include file="include/header.asp" -->
  <tr>
    <td class="first_content"><table cellpadding="5" cellspacing="0" border="0">
        <tr>
          <td valign="top"><img src="images/icon_stockmod.jpg" border="0" alt="Stock Modification" /></td>
          <td valign="top"><div class="alert alert-success"><img src="images/add_icon.png" border="0" align="bottom" /> <a href="add_stockmod.asp">Add Stock Modification</a></div></td>
          <td valign="top"><div class="alert alert-search">
              <form name="frmSearch" id="frmSearch" action="list_stockmod.asp?type=search" method="post" onsubmit="searchStockmod()">
                <h3>Search Parameters:</h3>
                Model Name / Part no / Model no / Created by:
                <input type="text" name="txtSearch" size="25" value="<%= request("txtSearch") %>" maxlength="20" />
                <select name="cboType" onchange="searchStockmod()">
                  <option value="">All Types</option>
                  <option <% if session("stockmod_type") = "KIT" then Response.Write " selected" end if%> value="KIT">Kit</option>
                  <option <% if session("stockmod_type") = "LEAD" then Response.Write " selected" end if%> value="LEAD">Lead</option>
                  <option <% if session("stockmod_type") = "PLUG" then Response.Write " selected" end if%> value="PLUG">Plug</option>
                  <option <% if session("stockmod_type") = "ADAPTOR" then Response.Write " selected" end if%> value="ADAPTOR">Adaptor</option>
                  <option <% if session("stockmod_type") = "ADAPTOR and DVD" then Response.Write " selected" end if%> value="ADAPTOR and DVD">Adaptor & DVD</option>
                  <option <% if session("stockmod_type") = "ADAPTOR and LEAD" then Response.Write " selected" end if%> value="ADAPTOR and LEAD">Adaptor & Lead</option>
                </select>
                <select name="cboSort" onchange="searchStockmod()">
                  <option <% if session("stockmod_sort") = "model_name" then Response.Write " selected" end if%> value="model_name">Sort by: Item name</option>
                  <option <% if session("stockmod_sort") = "model_type" then Response.Write " selected" end if%> value="model_type">Sort by: Type</option>
                  <option <% if session("stockmod_sort") = "part_no_base" then Response.Write " selected" end if%> value="part_no_base">Sort by: Part no</option>
                  <option <% if session("stockmod_sort") = "vendor_model_no" then Response.Write " selected" end if%> value="vendor_model_no">Sort by: Vendor model no</option>
                  <option <% if session("stockmod_sort") = "date_modified DESC" then Response.Write " selected" end if%> value="date_modified DESC">Sort by: Last modified</option>
                  <option <% if session("stockmod_sort") = "date_created DESC" then Response.Write " selected" end if%> value="date_created DESC">Sort by: Date created</option>
                </select>
                <input type="button" name="btnSearch" value="Search" onclick="searchStockmod()" />
                <input type="button" name="btnReset" value="Reset" onclick="resetSearch()" />
              </form>
            </div></td>
        </tr>
      </table>
      <table cellspacing="0" cellpadding="8" class="database_records_nowidth" width="1200">
        <thead>
          <tr>
            <td>Model no</td>
            <td>Takt Timing</td>
            <td>Type</td>
            <td>Part no BASE</td>
            <td>Vendor model no</td>
            <td>Document</td>
            <td>Created</td>
            <td>Last modified</td>
            <td>&nbsp;</td>
          </tr>
        </thead>
        <tbody>
          <%= strDisplayList %>
        </tbody>
      </table></td>
  </tr>
</table>
</body>
</html>