<!--#include file="include/connection_it.asp " -->
<% strSection = "stock_modification" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Add Stock Modification</title>
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
    reason += validateEmptyField(theForm.txtPartNoBase,"Part No BASE");
    reason += validateSpecialCharacters(theForm.txtPartNoBase,"Part No BASE");
    reason += validateSpecialCharacters(theForm.txtVendorModelNo,"Vendor Model No");

    if (reason != "") {
        alert("Some fields need correction:\n" + reason);

        blnSubmit = false;
        return false;
    }

    if (blnSubmit == true) {
        theForm.Action.value = 'Add';

        return true;
    }
}
</script>
<%

sub addStockMod
    dim strSQL
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

    call OpenDataBase()

    strSQL = "INSERT INTO yma_stock_mod ("
    strSQL = strSQL & "model_name, "
    strSQL = strSQL & "model_type, "
    strSQL = strSQL & "takt_timing, "
    strSQL = strSQL & "component1, qty1, "
    strSQL = strSQL & "component2, qty2, "
    strSQL = strSQL & "component3, qty3, "
    strSQL = strSQL & "component4, qty4, "
    strSQL = strSQL & "component5, qty5, "
    strSQL = strSQL & "component6, qty6, "
    strSQL = strSQL & "component7, qty7, "
    strSQL = strSQL & "component8, qty8, "
    strSQL = strSQL & "component9, qty9, "
    strSQL = strSQL & "component10,qty10,"
    strSQL = strSQL & "part_no_base, "
    strSQL = strSQL & "vendor_model_no, "
    strSQL = strSQL & "document, "
    strSQL = strSQL & "created_by, "
    strSQL = strSQL & "comments) VALUES ( "
    strSQL = strSQL & "'" & strModelName & "',"
    strSQL = strSQL & "'" & strModelType & "',"
    strSQL = strSQL & "'" & strTaktTiming & "',"
    strSQL = strSQL & "'" & strComponent1 & "', '" & strQty1 & "',"
    strSQL = strSQL & "'" & strComponent2 & "', '" & strQty2 & "',"
    strSQL = strSQL & "'" & strComponent3 & "', '" & strQty3 & "',"
    strSQL = strSQL & "'" & strComponent4 & "', '" & strQty4 & "',"
    strSQL = strSQL & "'" & strComponent5 & "', '" & strQty5 & "',"
    strSQL = strSQL & "'" & strComponent6 & "', '" & strQty6 & "',"
    strSQL = strSQL & "'" & strComponent7 & "', '" & strQty7 & "',"
    strSQL = strSQL & "'" & strComponent8 & "', '" & strQty8 & "',"
    strSQL = strSQL & "'" & strComponent9 & "', '" & strQty9 & "',"
    strSQL = strSQL & "'" & strComponent10 & "','" & strQty10 & "',"
    strSQL = strSQL & "'" & strPartNoBase & "',"
    strSQL = strSQL & "'" & strVendorModelNo & "',"
    strSQL = strSQL & "'" & intDocument & "',"
    strSQL = strSQL & "'" & session("UsrUserName") & "',"
    strSQL = strSQL & "'" & Server.HTMLEncode(strComments) & "')"

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
        emailSubject = "New Stock Mod: " & strModelName & " (" & strModelType &  ") by: " & session("UsrUserName")

        emailBodyText = "Created by: " & session("UsrUserName") & vbCrLf _
                      & "---------------------------------------------------------------------------" & vbCrLf _
                      & "NEW STOCK MOD DETAILS" & vbCrLf _
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

        Response.Redirect("thank-you_stockmod.asp")
    end if

    call CloseDataBase()
end sub

sub main
    call UTL_validateLogin

    if Request.ServerVariables("REQUEST_METHOD") = "POST" then
        if Trim(Request("Action")) = "Add" then
            call addStockMod
        end if
    end if
end sub

call main
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
            <h2>Add Stock Modification</h2>
            <font color="green"><%= strMessageText %></font></td>
        </tr>
      </table>
      <form action="" method="post" name="form_add_stockmod" id="form_add_stockmod" onsubmit="return validateFormOnSubmit(this)">
        <table cellpadding="5" cellspacing="0" class="item_maintenance_box">
          <tr>
            <td width="33%">Item name<span class="mandatory">*</span>:<br />
              <input type="text" id="txtModelName" name="txtModelName" maxlength="20" size="20" /></td>
            <td width="33%">Type:<br />
              <select name="cboModelType">
                <option value="-" rel="none">...</option>
                <option value="KIT" rel="component">KIT</option>
                <option value="LEAD" rel="none">LEAD</option>
                <option value="PLUG" rel="none">PLUG</option>
                <option value="ADAPTOR" rel="none">ADAPTOR</option>
                <option value="ADAPTOR and DVD" rel="none">ADAPTOR and DVD</option>
                <option value="ADAPTOR and LEAD" rel="none">ADAPTOR and LEAD</option>
              </select></td>
            <td width="33%">Takt Timing:<br />
                <input type="text" id="txtTaktTiming" name="txtTaktTiming" maxlength="20" size="20" /></td>
          </tr>
          <tr rel="component">
            <td colspan="3">Components:
              <table border="0">
                <tr>
                  <td>1.
                    <input type="text" id="txtComponent1" name="txtComponent1" maxlength="20" size="20" />
                    <input type="text" id="txtQty1" name="txtQty1" maxlength="3" size="5" /></td>
                </tr>
                <tr>
                  <td> 2.
                    <input type="text" id="txtComponent2" name="txtComponent2" maxlength="20" size="20" />
                    <input type="text" id="txtQty2" name="txtQty2" maxlength="3" size="5" /></td>
                </tr>
                <tr>
                  <td> 3.
                    <input type="text" id="txtComponent3" name="txtComponent3" maxlength="20" size="20" />
                    <input type="text" id="txtQty3" name="txtQty3" maxlength="3" size="5" /></td>
                </tr>
                <tr>
                  <td> 4.
                    <input type="text" id="txtComponent4" name="txtComponent4" maxlength="20" size="20" />
                    <input type="text" id="txtQty4" name="txtQty4" maxlength="3" size="5" /></td>
                </tr>
                <tr>
                  <td> 5.
                    <input type="text" id="txtComponent5" name="txtComponent5" maxlength="20" size="20" />
                    <input type="text" id="txtQty5" name="txtQty5" maxlength="3" size="5" /></td>
                </tr>
                <tr>
                  <td> 6.
                    <input type="text" id="txtComponent6" name="txtComponent6" maxlength="20" size="20" />
                    <input type="text" id="txtQty6" name="txtQty6" maxlength="3" size="5" /></td>
                </tr>
                <tr>
                  <td> 7.
                    <input type="text" id="txtComponent7" name="txtComponent7" maxlength="20" size="20" />
                    <input type="text" id="txtQty7" name="txtQty7" maxlength="3" size="5" /></td>
                </tr>
                <tr>
                  <td> 8.
                    <input type="text" id="txtComponent8" name="txtComponent8" maxlength="20" size="20" />
                    <input type="text" id="txtQty8" name="txtQty8" maxlength="3" size="5" /></td>
                </tr>
                <tr>
                  <td> 9.
                    <input type="text" id="txtComponent9" name="txtComponent9" maxlength="20" size="20" />
                    <input type="text" id="txtQty9" name="txtQty9" maxlength="3" size="5" /></td>
                </tr>
                <tr>
                  <td> 10.
                    <input type="text" id="txtComponent10" name="txtComponent10" maxlength="20" size="20" />
                    <input type="text" id="txtQty10" name="txtQty10" maxlength="3" size="5" /></td>
                </tr>
              </table></td>
          </tr>
          <tr>
            <td>Part no BASE<span class="mandatory">*</span>:<br />
              <input type="text" id="txtPartNoBase" name="txtPartNoBase" maxlength="20" size="20" /></td>
            <td colspan="2">Vendor Model No:<br />
              <input type="text" id="txtVendorModelNo" name="txtVendorModelNo" maxlength="20" size="20" /></td>
          </tr>
          <tr>
            <td width="50%" colspan="3"><input type="checkbox" name="chkDocument" id="chkDocument" value="1" /> Document</td>
          </tr>
          <tr>
            <td colspan="3">Comments:<br />
              <textarea name="txtComments" id="txtComments" cols="45" rows="3"></textarea></td>
          </tr>
        </table>
        <p>
          <input type="hidden" name="Action" />
          <input type="submit" value="Add Stock Modification" />
        </p>
      </form></td>
  </tr>
</table>
</body>
</html>