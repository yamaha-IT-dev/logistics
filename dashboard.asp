<% response.cookies("current_URL_cookie_logistics") = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?" & Request.Querystring %>
<!--#include file="include/connection_it.asp " -->
<% strSection = "dashboard" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Dashboard</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="include/autoSum2.js"></script>

<script language="JavaScript" type="text/javascript">
function baseModelLookup() {
    var strBase   = document.forms[0].txtBase.value;
    var strModel  = document.forms[0].txtModel.value;

    document.location.href = 'dashboard.asp?type=basemodellookup&base=' + strBase + '&model=' + strModel;
}

function resetLookup() {
    document.location.href = 'dashboard.asp?type=reset';
}
</script>
</head>
<body>
<%
Sub setLookup()
    Select Case Trim(Request("type"))
        case "reset"
            session("base") = ""
            session("model") = ""
        case "basemodellookup"
            session("base") = Trim(Request("base"))
            session("model") = Trim(Request("model"))
    End Select
End Sub

Sub displayBaseModelLookup()
    Dim base
    Dim model

    base = session("base")
    model = session("model")

    'we only expect one value as this is a lookup
    If base <> "" and model <> "" Then
        Response.Write("<script language=""javascript"">alert('Please provide either Base or Model code to lookup. Not both.');</script>")
    ElseIf (base = "" and model <> "") or (base <> "" and model = "") Then
        'lookup model
        If (base <> "") Then
            Call OpenBaseDataBase()

            Set rs = Server.CreateObject("ADODB.recordset")

            rs.CursorLocation = 3   'adUseClient
            rs.CursorType     = 3   'adOpenStatic
            rs.PageSize       = 100

            strSQL = "SELECT Y3MDBM FROM YF3MP WHERE Y3SOSC = '" & base & "'" & " AND Y3LFCY = '' AND Y3SKKI <> 'D'"

            rs.Open strSQL, conn

            If rs.EOF Then
                Response.Write("<script language=""javascript"">alert('No matching Model code found.');</script>")
            Else
                session("model") = rs("Y3MDBM")
            End If

            Call CloseBaseDataBase()
        End If

        'lookup base
        If (model <> "") Then
            Call OpenBaseDataBase()

            Set rs = Server.CreateObject("ADODB.recordset")

            rs.CursorLocation = 3   'adUseClient
            rs.CursorType     = 3   'adOpenStatic
            rs.PageSize       = 100

            strSQL = "SELECT Y3SOSC FROM YF3MP WHERE Y3MDBM = '" & model & "'" & " AND Y3LFCY = '' AND Y3SKKI <> 'D'"

            rs.Open strSQL, conn

            If rs.EOF Then
                Response.Write("<script language=""javascript"">alert('No matching Base code found.');</script>")
            Else
                session("base") = rs("Y3SOSC")
            End If

            Call CloseBaseDataBase()
        End If
    ElseIf base = "" and model = "" Then
        Response.Write("<script language=""javascript"">alert('Please provide either Base or Model code to lookup.');</script>")
    End If
End Sub

Sub main()
    If Trim(Request("type")) <> "" Then
        Call setLookup

        If Trim(Request("type")) = "basemodellookup" Then
        Call displayBaseModelLookup
        End If
    End If
End Sub

Call main
%>
<table cellspacing="0" cellpadding="0" align="center" class="full_size_table" border="0">
    <!-- #include file="include/header.asp" -->
    <tr>
        <td class="first_content" valign="top">
            <h1>Dashboard</h1>

            <%
            If Session("UsrLoginRole") <> 17 Then
            %>
            <table width="500" border="0" cellspacing="0" cellpadding="4" class="thin_border_grey">
                <tr>
                    <td>
                        <h2>Freight Reports</h2>
                        <ul>
                            <li><a href="../Divisions/Logistics/LogisticsIntraDocs/StarTrack_Datafiles/Report301.xls">Report 301</a></li>
                            <li><a href="../Divisions/Logistics/LogisticsIntraDocs/StarTrack_Datafiles/DeliveryTracker.xls">Delivery Tracker</a></li>
                            <li><a href="\\ymafp001\as400\SQLOutput\itemavail.xls">NZ Item Availability</a></li>
                            <li><a href="\\ymafp001\as400\SQLOutput\itemmaster.xls">Item Master</a></li>
                            <li><a href="\\ymafp001\as400\SQLOutput\newcode.xls">New Item Code</a></li>
                            <li><a href="\\ymafp001\as400\SQLOutput\ForwardOrders.xls">Forward Orders (hourly update)</a></li>
                            <li><a href="\\ymafp001\as400\SQLOutput\StockIntegrity.xls">BASE and TTL Stock Reconciliation</a></li>
                            <li><a href="\\ymafp001\as400\SQLOutput\nzorderm3.xls">NZ Order M3 Report</a></li>
                            <li><a href="\\ymafp001\as400\SQLOutput\Purchase Invoice Error List.xlsx">Purchase Invoice Error List</a></li>
                        </ul>
                    </td>
                </tr>
            </table>
            <br />
            <%
            End If
            %>

            <%
            If Session("UsrLoginRole") <> 17 Then
            %>
            <table width="500" border="0" cellspacing="0" cellpadding="4" class="thin_border_grey">
                <tr>
                    <td>
                        <h2>Stocktake</h2>
                        <ul>
                            <li><a href="file:\\YAMMAS22\accsdata$\Front End DB's\StockTake.accdb">Stocktake Link</a></li>
                        </ul>
                    </td>
                </tr>
            </table>
            <br />
            <%
            End If
            %>

            <table width="500" border="0" cellspacing="0" cellpadding="4" class="thin_border_grey">
                <tr>
                    <td>
                        <h2>Base to Model Lookup</h2>
                        <form name="frmBaseModelLookup" id="frmBaseModelLookup" action="dashboard.asp?type=basemodellookup" method="post" onsubmit="baseModelLookup()">
                            <input type="text" id="txtBase" name="txtBase" placeholder="Base" size="20" value="<%= session("base") %>" />
                            &harr;
                            <input type="text" id="txtModel" name="txtModel" placeholder="Model" size="20" value="<%= session("model") %>" />
                            <input type="button" id="btnLookup" name="btnLookup" value="Lookup" onclick="baseModelLookup()" />
                            <input type="button" name="btnLookupReset" value="Reset" onclick="resetLookup()" />
                        </form>
                    </td>
                </tr>
            </table>
            <br />

            <form action="" method="post" name="form_calculator" id="form_calculator">
                <table width="500" border="0" cellspacing="0" cellpadding="4" class="thin_border_grey">
                    <tr>
                        <td colspan="4"><h2>Calculator</h2></td>
                    </tr>
                    <tr>
                        <td><strong>H:</strong></td>
                        <td><strong>W:</strong></td>
                        <td><strong>D:</strong></td>
                        <td><strong>Cubic:</strong></td>
                    </tr>
                    <tr>
                        <td><input type="text" id="txtHeight" name="txtHeight" maxlength="6" size="8" onfocus="startCalc();" onblur="stopCalc();" /> cm</td>
                        <td><input type="text" id="txtWidth" name="txtWidth" maxlength="6" size="8" onfocus="startCalc();" onblur="stopCalc();" /> cm</td>
                        <td><input type="text" id="txtDepth" name="txtDepth" maxlength="6" size="8" onfocus="startCalc();" onblur="stopCalc();" /> cm</td>
                        <td><input type="text" id="txtTotal" name="txtTotal" maxlength="6" size="8" style="background-color:#CCC" /> m<sup>3</sup></td>
                    </tr>
                </table>
            </form>
        </td>
    </tr>
</table>

<script language="JavaScript" type="text/javascript">
    var base = document.getElementById('txtBase').value;
    var model = document.getElementById('txtModel').value;

    if (base != "" & model != "") {
        document.getElementById('txtBase').setAttribute("disabled", "disabled")
        document.getElementById('txtModel').setAttribute("disabled", "disabled")
        document.getElementById('btnLookup').setAttribute("disabled", "disabled")
    }
</script>

</body>
</html>