<% response.cookies("current_URL_cookie_logistics") = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?" & Request.Querystring %>
<!--#include file="include/connection_it.asp " -->
<!--#include file="include/listFolder.asp " -->
<!--#include file="class/clsComment.asp " -->
<!--#include file="class/cls3thReturn.asp " -->
<!--#include file="class/clsWarehouseReturn.asp " -->
<% strSection = "3TH" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Update 3TH</title>
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<link href="uploadify214/css/default.css" rel="stylesheet" type="text/css" />
<link href="uploadify214/css/uploadify.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="include/generic_form_validations.js"></script>
<script type="text/javascript" src="include/usableforms.js"></script>
<script type="text/javascript" src="include/javascript.js"></script>
<script language="JavaScript" type="text/javascript">
function validateFormOnSubmit(theForm) {
    var reason      = "";
    var blnSubmit   = true;

    reason += validateEmptyField(theForm.cboReturnType,"Type");

    reason += validateEmptyField(theForm.txtItemCode,"Item Code");
    reason += validateSpecialCharacters(theForm.txtItemCode,"Item Code");

    reason += validateEmptyField(theForm.txtDescription,"Item Description");
    reason += validateSpecialCharacters(theForm.txtDescription,"Item Description");

    reason += validateSpecialCharacters(theForm.txtDealer,"Dealer");

    reason += validateEmptyField(theForm.txtShipmentNo,"Shipment No");
    reason += validateSpecialCharacters(theForm.txtShipmentNo,"Shipment No");

    reason += validateNumeric(theForm.txtQty,"Qty");

    if (reason != "") {
        alert("Some fields need correction:\n" + reason);

        blnSubmit = false;
        return false;
    }

    if (blnSubmit == true) {
        theForm.Action.value = 'Update';

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
<script type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jquery/1.5/jquery.min.js"></script>
<script type="text/javascript" src="uploadify214/swfobject.js"></script>
<script type="text/javascript" src="uploadify214/jquery.uploadify.v2.1.4.min.js"></script>
<script type="text/javascript">
function Send_document() {
    $('#uploadify').uploadifyUpload();
}
</script>
<script type="text/javascript">
$(document).ready(function() {
    $("#uploadify").uploadify({
        'uploader'      : 'uploadify214/uploadify.swf',
        'script'        : 'uploader214.asp?sId=<%=session.sessionID%>',
        'cancelImg'     : 'uploadify214/cancel.png',
        'fileDesc'      : 'JPG (*.jpg), JPEG (*.jpeg), JPE (*.jpe), JP2 (*.jp2), JFIF (*.jfif), GIF (*.gif), BMP (*.bmp), PNG (*.png), PSD (*.psd), EPS (*.eps), ICO (*.ico), TIF (*.tif), TIFF (*.tiff), AI (*.ai), RAW (*.raw), TGA (*.tga), MNG (*.mng), SVG (*.svg), DOC (*.doc), RTF (*.rtf), TXT (*.txt), WPD (*.wpd), WPS (*.wps), CSV (*.csv), XML (*.xml), XSD (*.xsd), SQL (*.sql), PDF (*.pdf), XLS (*.xls), MDB (*.mdb), PPT (*.ppt), DOCX (*.docx), XLSX (*.xlsx), PPTX (*.pptx), PPSX (*.ppsx), ARTX (*.artx), MP3 (*.mp3), WMA (*.wma), MID (*.mid), MIDI (*.midi), MP4 (*.mp4), MPG (*.mpg), MPEG (*.mpeg), WAV (*.wav), RAM (*.ram), RA (*.ra), AVI (*.avi), MOV (*.mov), FLV (*.flv), M4A (*.m4a), M4V (*.m4v), HTM (*.htm), HTML (*.html), CSS (*.css), SWF (*.swf), JS (*.js), RAR (*.rar), ZIP (*.zip), TAR (*.tar), GZ (*.gz)',
        'fileExt'       : '*.jpg;*.jpeg;*.jpe;*.jp2;*.jfif;*.gif;*.bmp;*.png;*.psd;*.eps;*.ico;*.tif;*.tiff;*.ai;*.raw;*.tga;*.mng;*.svg;*.doc;*.rtf;*.txt;*.wpd;*.wps;*.csv;*.xml;*.xsd;*.sql;*.pdf;*.xls;*.mdb;*.ppt;*.docx;*.xlsx;*.pptx;*.ppsx;*.artx;*.mp3;*.wma;*.mid;*.midi;*.mp4;*.mpg;*.mpeg;*.wav;*.ram;*.ra;*.avi;*.mov;*.flv;*.m4a;*.m4v;*.htm;*.html;*.css;*.swf;*.js;*.rar;*.zip;*.tar;*.gz',
        'folder'        : '<%=application("uploadpath")%>',
        'multi'         : true,
        onError         : function (a, b, c, d) {
                            if (d.status == 404)
                                alert('Could not find upload script. Use a path relative to: '+'<?= getcwd() ?>');
                            else if (d.type === "HTTP")
                                alert('error ' + d.type + ": " + d.status);
                            else if (d.type === "File Size")
                                alert(c.name + ' ' + d.type + ' Limit: ' + Math.round(d.sizeLimit/1024) + 'KB');
                            else
                                alert('error ' + d.type + ": " + d.text);
                          },
        onComplete      : function(event, queueID, fileObj, response, data) {
                            var path = escape(fileObj.filePath);
                            $('#filesUploaded').append('<div class=\'uploadifyQueueItem\'><a href='+path+' target=\'_blank\'>'+fileObj.name+'</a></div>');
                            location.reload();
                          }
    });
});
</script>

<%
application("sessionID")  = Session.SessionID
application("uploadpath") = "3th/" & request("id") & ""

sub main
    call UTL_validateLogin

    dim intReturnID
    intReturnID = request("id")

    if Request.ServerVariables("REQUEST_METHOD") = "POST" then
        dim intReturnType
        dim strDepartment
        dim strItemCode
        dim strShipmentNo
        dim intQty
        dim strDescription
        dim strGRA
        dim strCarrier
        dim strLabelNo
        dim strOriginalConnote
        dim strDealer
        dim intInstruction
        dim strSerialNo
        dim strDateReceived
        dim strComments
        dim intStatus

        intReturnType       = Request.Form("cboReturnType")
        strDepartment       = Trim(Request.Form("cboDepartment"))
        strItemCode         = Replace(Trim(Request.Form("txtItemCode")),"'","''")
        strShipmentNo       = Replace(Trim(Request.Form("txtShipmentNo")),"'","''")
        intQty              = Request.Form("txtQty")
        strDescription      = Replace(Trim(Request.Form("txtDescription")),"'","''")
        strGRA              = Replace(Trim(Request.Form("txtGRA")),"'","''")
        strCarrier          = Request.Form("cboCarrier")
        strLabelNo          = Replace(Trim(Request.Form("txtLabelNo")),"'","''")
        strOriginalConnote  = Replace(Trim(Request.Form("txtOriginalConnote")),"'","''")
        strDealer           = Replace(Trim(Request.Form("txtDealer")),"'","''")
        intInstruction      = Request.Form("cboInstruction")
        strSerialNo         = Replace(Trim(Request.Form("txtSerialNo")),"'","''")
        strDateReceived     = Replace(Trim(Request.Form("txtDateReceived")),"'","''")
        strComments         = Replace(Trim(Request.Form("txtComments")),"'","''")
        intStatus           = Request.Form("cboStatus")

        select case Trim(Request("Action"))
            case "Update"
                call update3thReturn(intReturnID, intReturnType, strDepartment, strItemCode, strShipmentNo, intQty, strDescription, strGRA, strCarrier, strLabelNo, strOriginalConnote, strDealer, intInstruction, strSerialNo, strDateReceived, strComments, intStatus, session("UsrUserName"))	
            case "Comment"
                call addComment(intReturnID,warehouse3thReturnModuleID)
        end select
    end if

    call get3thReturn(intReturnID)
    call getReasonCode
    call listComments(intReturnID,warehouse3thReturnModuleID)
end sub

call main

dim strMessageText
dim strCommentsList
dim strReasonCodeList
%>
</head>
<body>
<table cellspacing="0" cellpadding="0" align="center" class="full_size_table" border="0">
  <!-- #include file="include/header.asp" -->
  <tr>
    <td class="first_content">
      <table cellpadding="5" cellspacing="0" border="0">
        <tr>
          <td><a href="list_3TH.asp"><img src="images/icon_return.jpg" border="0" alt="Warehouse Return" /></a></td>
          <td valign="top">
            <img src="images/backward_arrow.gif" border="0" /> <a href="list_3TH.asp">Back to List</a>
            <h2>Update 3TH</h2>
          </td>
          <td valign="top">
            <table cellpadding="4" cellspacing="0" class="created_table">
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
            </table>
          </td>
        </tr>
      </table>
      <table>
        <tr>
          <td>
            <form action="" method="post" name="form_update_quarantine" id="form_update_quarantine" onsubmit="return validateFormOnSubmit(this)">
              <table cellpadding="5" cellspacing="0" class="item_maintenance_box">
                <tr>
                  <td colspan="3" align="center" class="item_maintenance_header">
                    Lapsed: <%= session("days_in_3TH") %> days
                  </td>
                </tr>
                <tr>
                  <td colspan="3">
                    Type<span class="mandatory">*</span>:
                    <select name="cboReturnType">
                      <option <% if session("return_type") = "" then Response.Write " selected" end if%> value="">...</option>
                      <option <% if session("return_type") = "1" then Response.Write " selected" end if%> value="1">Lost in Warehouse</option>
                      <option <% if session("return_type") = "2" then Response.Write " selected" end if%> value="2">Lost by Carrier</option>
                      <option <% if session("return_type") = "3" then Response.Write " selected" end if%> value="3">Packaging Issue</option>
                      <option <% if session("return_type") = "4" then Response.Write " selected" end if%> value="4">Warehouse Variance</option>
                      <option <% if session("return_type") = "5" then Response.Write " selected" end if%> value="5">Display Stock</option>
                    </select>
                  </td>
                </tr>
                <tr>
                  <td>
                    Department<span class="mandatory">*</span>:
                    <br />
                    <select name="cboDepartment">
                      <option <% if Trim(session("3thDepartment")) = "AV" then Response.Write " selected" end if%> value="AV">AV</option>
                      <option <% if Trim(session("3thDepartment")) = "MPD" then Response.Write " selected" end if%> value="MPD">MPD</option>
                      <option <% if Trim(session("3thDepartment")) = "Other" then Response.Write " selected" end if%> value="Other">Other</option>
                    </select>
                  </td>
                  <td colspan="2">
                    Item code<span class="mandatory">*</span>:
                    <br />
                    <input type="text" id="txtItemCode" name="txtItemCode" maxlength="20" size="30" value="<%= Server.HTMLEncode(session("item_code")) %>" />
                  </td>
                </tr>
                <tr>
                  <td colspan="3">
                    Item description<span class="mandatory">*</span>:
                    <br />
                    <input type="text" id="txtDescription" name="txtDescription" maxlength="50" size="60" value="<%= session("description") %>" />
                  </td>
                </tr>
                <tr>
                  <td>
                    Label #:
                    <br />
                    <input type="text" id="txtLabelNo" name="txtLabelNo" maxlength="15" size="20" value="<%= session("label_no") %>" />
                  </td>
                  <td colspan="2">
                    Con-note:<br />
                    <input type="text" id="txtOriginalConnote" name="txtOriginalConnote" maxlength="15" size="20" value="<%= session("original_connote") %>" />
                  </td>
                </tr>
                <tr>
                  <td width="50%">
                    Dealer:
                    <br />
                    <input type="text" id="txtDealer" name="txtDealer" maxlength="30" size="35" value="<%= session("dealer") %>" />
                  </td>
                  <td width="30%">
                    Shipment #<span class="mandatory">*</span>:
                    <br />
                    <input type="text" id="txtShipmentNo" name="txtShipmentNo" maxlength="10" size="12" value="<%= session("shipment_no") %>" />
                  </td>
                  <td width="20%">
                    Qty<span class="mandatory">*</span>:
                    <br />
                    <input type="text" id="txtQty" name="txtQty" maxlength="3" size="4" value="<%= session("qty") %>" />
                  </td>
                </tr>
                <tr>
                  <td>
                    Carrier:
                    <br />
                    <select name="cboCarrier">
                      <option <% if session("carrier") = "" then Response.Write " selected" end if%> value="">...</option>
                      <option <% if session("carrier") = "Cope" then Response.Write " selected" end if%> value="Cope">Cope</option>
                      <option <% if session("carrier") = "StarTrack" then Response.Write " selected" end if%> value="StarTrack">StarTrack</option>
                      <option <% if session("carrier") = "Schenker" then Response.Write " selected" end if%> value="Schenker">Schenker</option>
                      <option <% if session("carrier") = "Kings" then Response.Write " selected" end if%> value="Kings">Kings</option>
                    </select>
                  </td>
                  <td colspan="2">
                    Date received:
                    <br />
                    <input type="text" id="txtDateReceived" name="txtDateReceived" maxlength="10" size="15" value="<%= session("date_received") %>" />
                  </td>
                </tr>
                <tr>
                  <td colspan="3">
                    Serial #:
                    <br />
                    <input type="text" id="txtSerialNo" name="txtSerialNo" maxlength="50" size="55" value="<%= session("serial_no") %>" />
                  </td>
                </tr>
                <tr>
                  <td>
                    Instruction:
                    <br />
                    <select name="cboInstruction">
                      <option <% if session("instruction") = "" then Response.Write " selected" end if%> value="">...</option>
                      <option <% if session("instruction") = "1" then Response.Write " selected" end if%> value="1">Update GRA</option>
                      <option <% if session("instruction") = "2" then Response.Write " selected" end if%> value="2">Writeoff Approval Required</option>
                    </select>
                  </td>
                  <td colspan="2">
                    GRA:
                    <br />
                    <input type="text" id="txtGRA" name="txtGRA" maxlength="7" size="10" value="<%= session("gra") %>" />
                  </td>
                </tr>
                <tr>
                  <td colspan="3">
                    Comments:
                    <br />
                    <textarea name="txtComments" id="txtComments" cols="50" rows="4"><%= session("comments") %></textarea>
                  </td>
                </tr>
                <tr class="status_row">
                  <td colspan="3">
                    Status:
                    <select name="cboStatus">
                      <option <% if session("status") = "1" then Response.Write " selected" end if%> value="1">Open</option>
                      <option <% if session("status") = "0" then Response.Write " selected" end if%> value="0">Completed</option>
                    </select>
                  </td>
                </tr>
                <tr>
                  <td colspan="3">
                    <input type="hidden" name="Action" />
                    <input type="submit" value="Update 3TH" />
                  </td>
                </tr>
              </table>
            </form>
          </td>
          <td valign="top">
            <table cellpadding="5" cellspacing="0" class="serial_no_box">
              <tr>
                <td class="item_maintenance_header">Documents</td>
              </tr>
              <tr>
                <td>
                  <form id="formIDdoc" name="formIDdoc" class="form" method="post">
                    <p>Please select the files first, then click the "Upload" button.</p>
                    <p><input class="text-input" name="uploadify" id="uploadify" type="file" size="20" /></p>
                    <h3 align="right"><a href="javascript:Send_document()"><img src="images/btn_upload.gif" border="0" align="top" /></a></h3>
                    <div id="filesUploaded"></div>
                  </form>
                  <p><% ListFolderContents(Server.MapPath("3th/" & request("id") & "")) %></p>
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
      <%= strMessageText %>
      <h2>
        Comments
        <br />
        <img src="images/comment_bar.jpg" border="0" />
      </h2>
      <table cellpadding="5" cellspacing="0" border="0" class="comments_box">
        <%= strCommentsList %>
        <tr>
          <td>
            <form action="" name="form_add_comment" id="form_add_comment" method="post" onsubmit="return submitComment(this)">
              <p>
                <input type="text" name="txtComment" id="txtComment" maxlength="60" size="65" />
                <input type="hidden" name="Action" />
                <input type="submit" value="Add Comment" />
              </p>
            </form>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>

<script type="text/javascript" src="include/moment.js"></script>
<script type="text/javascript" src="include/pikaday.js"></script>

<script type="text/javascript">
    var picker = new Pikaday({
        field: document.getElementById('txtDateReceived'),
        firstDay: 1,
        minDate: new Date('1900-01-01'),
        maxDate: new Date('2020-12-31'),
        yearRange: [1900,2020],
        format: 'DD/MM/YYYY'
    });
</script>
</body>
</html>