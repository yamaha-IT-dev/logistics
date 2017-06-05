<% response.cookies("current_URL_cookie_logistics") = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?" & Request.Querystring %>
<!--#include file="include/connection_it.asp " -->
<!--#include file="class/clsComment.asp " -->
<% strSection = "stock_modification" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Update Stock Modification</title>
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="include/usableforms.js"></script>
<script type="text/javascript" src="include/generic_form_validations.js"></script>
<script language="JavaScript" type="text/javascript">
function validateFormOnSubmit(theForm) {
    var reason = "";
    var blnSubmit = true;

    reason += validateEmptyField(theForm.txtModelName,"Item name");
    reason += validateSpecialCharacters(theForm.txtModelName,"Item name");
    reason += validateSpecialCharacters(theForm.txtTaktTiming,"Takt timing");
    reason += validateEmptyField(theForm.txtPartNoBase,"Part no BASE");
    reason += validateSpecialCharacters(theForm.txtPartNoBase,"Part no BASE");
    reason += validateSpecialCharacters(theForm.txtVendorModelNo,"Vendor model no");

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
</script>
<%
Sub getStockMod
    dim intID
    intID = request("id")

    call OpenDataBase()

    set rs = Server.CreateObject("ADODB.recordset")

    rs.CursorLocation = 3   'adUseClient
    rs.CursorType = 3       'adOpenStatic

    strSQL = "SELECT * FROM yma_stock_mod WHERE stock_id = " & intID

    rs.Open strSQL, conn

    'Response.Write strSQL

    if not DB_RecSetIsEmpty(rs) Then
        session("model_name")       = rs("model_name")
        session("model_type")       = rs("model_type")
        session("takt_timing")      = rs("takt_timing")
        session("component1")       = rs("component1")
        session("qty1")             = rs("qty1")
        session("component2")       = rs("component2")
        session("qty2")             = rs("qty2")
        session("component3")       = rs("component3")
        session("qty3")             = rs("qty3")
        session("component4")       = rs("component4")
        session("qty4")             = rs("qty4")
        session("component5")       = rs("component5")
        session("qty5")             = rs("qty5")
        session("component6")       = rs("component6")
        session("qty6")             = rs("qty6")
        session("component7")       = rs("component7")
        session("qty7")             = rs("qty7")
        session("component8")       = rs("component8")
        session("qty8")             = rs("qty8")
        session("component9")       = rs("component9")
        session("qty9")             = rs("qty9")
        session("component10")      = rs("component10")
        session("qty10")            = rs("qty10")
        session("part_no_base")     = rs("part_no_base")
        session("vendor_model_no")  = rs("vendor_model_no")
        session("document")         = rs("document")
        session("status")           = rs("status")
        session("date_created")     = rs("date_created")
        session("created_by")       = rs("created_by")
        session("date_modified")    = rs("date_modified")
        session("modified_by")      = rs("modified_by")
        session("comments")         = rs("comments")
    end if

    call CloseDataBase()
end sub

sub updateStockMod
    dim strSQL
    dim intID
    intID = request("id")

    dim strModelName
    dim strModelType
    dim strTaktTiming
    dim strComponent1
    dim strQty1
    dim strComponent2
    dim strQty2
    dim strComponent3
    dim strQty3
    dim strComponent4
    dim strQty4
    dim strComponent5
    dim strQty5
    dim strComponent6
    dim strQty6
    dim strComponent7
    dim strQty7
    dim strComponent8
    dim strQty8
    dim strComponent9
    dim strQty9
    dim strComponent10
    dim strQty10
    dim strPartNoBase
    dim strVendorModelNo
    dim intDocument
    dim strComments
    dim intStatus

    strModelName        = Replace(Request.Form("txtModelName"),"'","''")
    strModelType        = Request.Form("cboModelType")
    strTaktTiming       = Replace(Request.Form("txtTaktTiming"),"'","''")
    strComponent1       = Replace(Request.Form("txtComponent1"),"'","''")
    strQty1             = Replace(Request.Form("txtQty1"),"'","''")
    strComponent2       = Replace(Request.Form("txtComponent2"),"'","''")
    strQty2             = Replace(Request.Form("txtQty2"),"'","''")
    strComponent3       = Replace(Request.Form("txtComponent3"),"'","''")
    strQty3             = Replace(Request.Form("txtQty3"),"'","''")
    strComponent4       = Replace(Request.Form("txtComponent4"),"'","''")
    strQty4             = Replace(Request.Form("txtQty4"),"'","''")
    strComponent5       = Replace(Request.Form("txtComponent5"),"'","''")
    strQty5             = Replace(Request.Form("txtQty5"),"'","''")
    strComponent6       = Replace(Request.Form("txtComponent6"),"'","''")
    strQty6             = Replace(Request.Form("txtQty6"),"'","''")
    strComponent7       = Replace(Request.Form("txtComponent7"),"'","''")
    strQty7             = Replace(Request.Form("txtQty7"),"'","''")
    strComponent8       = Replace(Request.Form("txtComponent8"),"'","''")
    strQty8             = Replace(Request.Form("txtQty8"),"'","''")
    strComponent9       = Replace(Request.Form("txtComponent9"),"'","''")
    strQty9             = Replace(Request.Form("txtQty9"),"'","''")
    strComponent10      = Replace(Request.Form("txtComponent10"),"'","''")
    strQty10            = Replace(Request.Form("txtQty10"),"'","''")
    strPartNoBase       = Replace(Request.Form("txtPartNoBase"),"'","''")
    strVendorModelNo    = Replace(Request.Form("txtVendorModelNo"),"'","''")
    intDocument         = Request.Form("chkDocument")
    strComments         = Replace(Request.Form("txtComments"),"'","''")
    intStatus           = Request.Form("cboStatus")

    Call OpenDataBase()

    strSQL = "UPDATE yma_stock_mod SET "
    strSQL = strSQL & "model_name = '" & strModelName & "',"
    strSQL = strSQL & "model_type = '" & strModelType & "',"
    strSQL = strSQL & "takt_timing = '" & strTaktTiming & "',"
    strSQL = strSQL & "component1 = '" & strComponent1 & "',"
    strSQL = strSQL & "qty1 = '" & strQty1 & "',"
    strSQL = strSQL & "component2 = '" & strComponent2 & "',"
    strSQL = strSQL & "qty2 = '" & strQty2 & "',"
    strSQL = strSQL & "component3 = '" & strComponent3 & "',"
    strSQL = strSQL & "qty3 = '" & strQty3 & "',"
    strSQL = strSQL & "component4 = '" & strComponent4 & "',"
    strSQL = strSQL & "qty4 = '" & strQty4 & "',"
    strSQL = strSQL & "component5 = '" & strComponent5 & "',"
    strSQL = strSQL & "qty5 = '" & strQty5 & "',"
    strSQL = strSQL & "component6 = '" & strComponent6 & "',"
    strSQL = strSQL & "qty6 = '" & strQty6 & "',"
    strSQL = strSQL & "component7 = '" & strComponent7 & "',"
    strSQL = strSQL & "qty7 = '" & strQty7 & "',"
    strSQL = strSQL & "component8 = '" & strComponent8 & "',"
    strSQL = strSQL & "qty8 = '" & strQty8 & "',"
    strSQL = strSQL & "component9 = '" & strComponent9 & "',"
    strSQL = strSQL & "qty9 = '" & strQty9 & "',"
    strSQL = strSQL & "component10 = '" & strComponent10 & "',"
    strSQL = strSQL & "qty10 = '" & strQty10 & "',"
    strSQL = strSQL & "part_no_base = '" & strPartNoBase & "',"
    strSQL = strSQL & "vendor_model_no = '" & strVendorModelNo & "',"
    strSQL = strSQL & "document = '" & intDocument & "',"
    strSQL = strSQL & "comments = '" & strComments & "',"
    strSQL = strSQL & "date_modified = getdate(),"
    strSQL = strSQL & "modified_by = '" & session("UsrUserName") & "',"
    strSQL = strSQL & "status = '" & intStatus & "' WHERE stock_id = " & intID

    'response.Write strSQL
    Set oMail = Server.CreateObject("CDO.Message")
    Set iConf = Server.CreateObject("CDO.Configuration")
    Set Flds = iConf.Fields

    iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.sendgrid.net"
    iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 10
    iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
    iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 'basic clear text
    iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "yamahamusicau"
    iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "str0ppy@16"
    iConf.Fields.Update

    emailFrom = "automailer@music.yamaha.com"

    on error resume next
    conn.Execute strSQL

    if err <> 0 then
        strMessageText = err.description
    else
        'emailTo = "harsono_setiono@gmx.yamaha.com"
        emailTo = "YMA-Warehouse@ttlogistics.com.au"
        emailSubject = "Updated Stock Mod: " & strModelName & " (" & strModelType &  ") by: " & session("UsrUserName")
        emailBodyText = "Updated by: " & session("UsrUserName") & vbCrLf _											
                      & "---------------------------------------------------------------------------" & vbCrLf _
                      & "UPDATED STOCK MOD DETAILS" & vbCrLf _
                      & "---------------------------------------------------------------------------" & vbCrLf _
                      & "Item name:       " & strModelName & vbCrLf _
                      & "Type:            " & strModelType & vbCrLf _
                      & "Takt Timing:     " & strTaktTiming & vbCrLf _
                      & "Part no (BASE):  " & strPartNoBase & vbCrLf _
                      & "Vendor model no: " & strVendorModelNo & vbCrLf _
                      & "Document:        " & intDocument & vbCrLf _
                      & "Comments:        " & strComments & vbCrLf _
                      & "---------------------------------------------------------------------------" & vbCrLf _
                      & " " & vbCrLf _
                      & "Please click on the below link for more info:" & vbCrLf _
                      & "http://intranet/logistics/update_stockmod.asp?id=" & intID & "" & vbCrLf _
                      & " " & vbCrLf _
                      & "This is an automated email - please do not reply to this email."

        Set oMail.Configuration = iConf
        oMail.To        = emailTo
        oMail.Cc        = emailCc
        oMail.Bcc       = emailBcc
        oMail.From      = emailFrom
        oMail.Subject   = emailSubject
        oMail.TextBody  = emailBodyText
        oMail.Send

        Set iConf = Nothing
        Set Flds = Nothing

        strMessageText = "The stock modification has been updated."
    end if

    Call CloseDataBase()
end sub

sub backRecordButton
    dim strSQL
    dim intID
    intID = request("id")

    dim strDisplayBackButton

    set rs = Server.CreateObject("ADODB.recordset")

    rs.CursorLocation = 3   'adUseClient
    rs.CursorType = 3       'adOpenStatic

    Call OpenDataBase()
    strSQL = "SELECT TOP 1 stock_id FROM yma_stock_mod WHERE stock_id < '" & intID & "' ORDER BY stock_id DESC"

    rs.Open strSQL, conn

    if not DB_RecSetIsEmpty(rs) Then
        session("previous_button") = "<a href=""update_stockmod.asp?id=" & rs("stock_id") & """><img src=""images/backpage.png"" border=""0"" /></a>"
    end if

    call CloseDataBase()
end sub

sub nextRecordButton
    dim strSQL
    dim intID
    intID = request("id")

    dim strDisplayNextButton

    set rs = Server.CreateObject("ADODB.recordset")

    rs.CursorLocation = 3   'adUseClient
    rs.CursorType = 3       'adOpenStatic

    Call OpenDataBase()
    strSQL = "SELECT TOP 1 stock_id FROM yma_stock_mod WHERE stock_id > '" & intID & "' ORDER BY stock_id"

    rs.Open strSQL, conn

    if not DB_RecSetIsEmpty(rs) Then
        session("next_button") = "<a href=""update_stockmod.asp?id=" & rs("stock_id") & """><img src=""images/nextpage.png"" border=""0"" /></a>"
    end if

    call CloseDataBase()
end sub

sub main
    call UTL_validateLogin

    dim intID
    intID = request("id")

    call getStockMod
    call listComments(intID,stockmodModuleID)

    'call backRecordButton
    'call nextRecordButton

    if Request.ServerVariables("REQUEST_METHOD") = "POST" then
        select case Trim(Request("Action"))
            case "Update"
                call updateStockMod
                call getStockMod
            case "Comment"
                call addComment(intID,stockmodModuleID)
                call listComments(intID,stockmodModuleID)
        end select
    end if
end sub

call main

dim strMessageText
dim strCommentsList
%>
</head>
<body>
<table cellspacing="0" cellpadding="0" align="center" class="full_size_table" border="0">
  <!-- #include file="include/header.asp" -->
  <tr>
    <td class="first_content"><table cellpadding="5" cellspacing="0" border="0">
        <tr>
          <td valign="top"><a href="list_stockmod.asp"><img src="images/icon_stockmod.jpg" border="0" alt="Stock Modification" /></a></td>
          <td valign="top"><img src="images/backward_arrow.gif" border="0" /> <a href="list_stockmod.asp">Back to List</a>
            <h2>Update Stock Modification</h2>
            <font color="green"><%= strMessageText %></font></td>
          <td valign="top"><table cellpadding="4" cellspacing="0" class="created_table">
                <tr>
                  <td class="created_column_1"><strong>Created:</strong></td>
                  <td class="created_column_2"><%= session("created_by") %></td>
                  <td class="created_column_3"><%= displayDateFormatted(session("date_created")) %></td>
                </tr>
                <tr>
                  <td><strong>Last modified:</strong></td>
                  <td><%= session("modified_by") %></td>
                  <td><%= displayDateFormatted(session("date_modified")) %></td>
                </tr>
              </table></td>
        </tr>
      </table>
      <form action="" method="post" name="form_update_stockmod" id="form_update_stockmod" onsubmit="return validateFormOnSubmit(this)">
        <table cellpadding="5" cellspacing="0" class="item_maintenance_box">
          <tr>
            <td width="33%">Item name<span class="mandatory">*</span>:<br />
              <input type="text" id="txtModelName" name="txtModelName" maxlength="20" size="20" value="<%= session("model_name") %>" /></td>
            <td width="33%">Type:<br />
              <select name="cboModelType">
                <option <% if session("model_type") = "-" then Response.Write " selected" end if%> value="-" rel="none">...</option>
                <option <% if session("model_type") = "KIT" then Response.Write " selected" end if%> value="KIT" rel="component">KIT</option>
                <option <% if session("model_type") = "LEAD" then Response.Write " selected" end if%> value="LEAD" rel="none">LEAD</option>
                <option <% if session("model_type") = "PLUG" then Response.Write " selected" end if%> value="PLUG" rel="none">PLUG</option>
                <option <% if session("model_type") = "ADAPTOR" then Response.Write " selected" end if%> value="ADAPTOR" rel="none">ADAPTOR</option>
                <option <% if session("model_type") = "ADAPTOR and DVD" then Response.Write " selected" end if%> value="ADAPTOR and DVD" rel="none">ADAPTOR and DVD</option>
                <option <% if session("model_type") = "ADAPTOR and LEAD" then Response.Write " selected" end if%> value="ADAPTOR and LEAD" rel="none">ADAPTOR and LEAD</option>
              </select></td>
            <td width="33%">Takt Timing<br />
                <input type="text" id="txtTaktTiming" name="txtTaktTiming" maxlength="20" size="20" value="<%= session("takt_timing") %>" /></td>
          </tr>
          <tr rel="component">
            <td colspan="3">Components:
<table border="0">
                <tr>
                  <td>1.
                    <input type="text" id="txtComponent1" name="txtComponent1" maxlength="20" size="20" value="<%= session("component1") %>" />
                    <input type="text" id="txtQty1" name="txtQty1" maxlength="3" size="5" value="<%= session("qty1") %>" /></td>
                </tr>
                <tr>
                  <td> 2.
                    <input type="text" id="txtComponent2" name="txtComponent2" maxlength="20" size="20" value="<%= session("component2") %>" />
                    <input type="text" id="txtQty2" name="txtQty2" maxlength="3" size="5" value="<%= session("qty2") %>" /></td>
                </tr>
                <tr>
                  <td> 3.
                    <input type="text" id="txtComponent3" name="txtComponent3" maxlength="20" size="20" value="<%= session("component3") %>" />
                    <input type="text" id="txtQty3" name="txtQty3" maxlength="3" size="5" value="<%= session("qty3") %>" /></td>
                </tr>
                <tr>
                  <td> 4.
                    <input type="text" id="txtComponent4" name="txtComponent4" maxlength="20" size="20" value="<%= session("component4") %>" />
                    <input type="text" id="txtQty4" name="txtQty4" maxlength="3" size="5" value="<%= session("qty4") %>" /></td>
                </tr>
                <tr>
                  <td> 5.
                    <input type="text" id="txtComponent5" name="txtComponent5" maxlength="20" size="20" value="<%= session("component5") %>" />
                    <input type="text" id="txtQty5" name="txtQty5" maxlength="3" size="5" value="<%= session("qty5") %>" /></td>
                </tr>
                <tr>
                  <td> 6.
                    <input type="text" id="txtComponent6" name="txtComponent6" maxlength="20" size="20" value="<%= session("component6") %>" />
                    <input type="text" id="txtQty6" name="txtQty6" maxlength="3" size="5" value="<%= session("qty6") %>" /></td>
                </tr>
                <tr>
                  <td> 7.
                    <input type="text" id="txtComponent7" name="txtComponent7" maxlength="20" size="20" value="<%= session("component7") %>" />
                    <input type="text" id="txtQty7" name="txtQty7" maxlength="3" size="5" value="<%= session("qty7") %>" /></td>
                </tr>
                <tr>
                  <td> 8.
                    <input type="text" id="txtComponent8" name="txtComponent8" maxlength="20" size="20" value="<%= session("component8") %>" />
                    <input type="text" id="txtQty8" name="txtQty8" maxlength="3" size="5" value="<%= session("qty8") %>" /></td>
                </tr>
                <tr>
                  <td> 9.
                    <input type="text" id="txtComponent9" name="txtComponent9" maxlength="20" size="20" value="<%= session("component9") %>" />
                    <input type="text" id="txtQty9" name="txtQty9" maxlength="3" size="5" value="<%= session("qty9") %>" /></td>
                </tr>
                <tr>
                  <td> 10.
                    <input type="text" id="txtComponent10" name="txtComponent10" maxlength="20" size="20" value="<%= session("component10") %>" />
                    <input type="text" id="txtQty10" name="txtQty10" maxlength="3" size="5" value="<%= session("qty10") %>" /></td>
                </tr>
              </table></td>
          </tr>
          <tr>
            <td colspan="2">Part no BASE<span class="mandatory">*</span>:<br />
              <input type="text" id="txtPartNoBase" name="txtPartNoBase" maxlength="20" size="20" value="<%= session("part_no_base") %>" /></td>
            <td>Vendor model no:<br />
              <input type="text" id="txtVendorModelNo" name="txtVendorModelNo" maxlength="20" size="20" value="<%= session("vendor_model_no") %>" /></td>
          </tr>
          <tr>
            <td width="50%" colspan="3"><input type="checkbox" name="chkDocument" id="chkDocument" value="1" <% if session("document") = "1" then Response.Write " checked" end if%> />
              Document</td>
          </tr>
          <tr>
            <td colspan="3">Comments:<br />
              <textarea name="txtComments" id="txtComments" cols="45" rows="3"><%= session("comments") %></textarea></td>
          </tr>
          <tr class="status_row">
            <td colspan="3">Status: <select name="cboStatus">
                <option <% if session("status") = "1" then Response.Write " selected" end if%> value="1">Open</option>
                <option <% if session("status") = "0" then Response.Write " selected" end if%> value="0">Completed</option>
            </select></td>
          </tr>
        </table>
        <p>
          <input type="hidden" name="Action" />
          <input type="submit" value="Update Stock Modification" />
        </p>
      </form>
      <h2>Comments<br />
        <img src="images/comment_bar.jpg" border="0" /></h2>
      <table cellpadding="5" cellspacing="0" border="0" class="comments_box">
        <%= strCommentsList %>
        <tr>
          <td><form action="" name="form_add_comment" id="form_add_comment" method="post" onsubmit="return submitComment(this)">
              <p><input type="text" name="txtComment" id="txtComment" maxlength="60" size="65" />
              <input type="hidden" name="Action" />
              <input type="submit" value="Add Comment" /></p>
            </form></td>
        </tr>
      </table>
      </td>
  </tr>
</table>
</body>
</html>