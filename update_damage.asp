<% response.cookies("current_URL_cookie_logistics") = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?" & Request.Querystring %>
<!--#include file="include/connection_it.asp " -->
<!--#include file="include/listFolder.asp " -->
<!--#include file="class/clsDamageType.asp" -->
<!--#include file="class/clsComment.asp " -->
<% strSection = "damage" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Update Damaged Item</title>
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<link href="uploadify214/css/default.css" rel="stylesheet" type="text/css" />
<link href="uploadify214/css/uploadify.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="include/usableforms.js"></script>
<script type="text/javascript" src="include/generic_form_validations.js"></script>
<script language="JavaScript" type="text/javascript">
function validateFormOnSubmit(theForm) {
	var reason 		= "";
	var blnSubmit 	= true;
	
	reason += validateEmptyField(theForm.txtItemName,"Item name");
	reason += validateEmptyField(theForm.txtSerialNo,"Serial no");	
	
  	if (reason != "") {
    	alert("Some fields need correction:\n" + reason);
    	
		blnSubmit = false;
		return false;
  	}

	if (blnSubmit == true){
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
	
	if (blnSubmit == true){
		theForm.Action.value = 'Comment';
		
		return true;		
    }
}
</script>
<script type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jquery/1.5/jquery.min.js"></script>
<script type="text/javascript" src="uploadify214/swfobject.js"></script>
<script type="text/javascript" src="uploadify214/jquery.uploadify.v2.1.4.min.js"></script>
<script type="text/javascript">
function Send_document()
{
	$('#uploadify').uploadifyUpload();
}
</script>
<script type="text/javascript">
$(document).ready(function() {
	$("#uploadify").uploadify({
		'uploader'       : 'uploadify214/uploadify.swf',
		'script'         : 'uploader214.asp?sId=<%=session.sessionID%>',
		'cancelImg'      : 'uploadify214/cancel.png',		
		'fileDesc'		 : 'JPG (*.jpg), JPEG (*.jpeg), JPE (*.jpe), JP2 (*.jp2), JFIF (*.jfif), GIF (*.gif), BMP (*.bmp), PNG (*.png), PSD (*.psd), EPS (*.eps), ICO (*.ico), TIF (*.tif), TIFF (*.tiff), AI (*.ai), RAW (*.raw), TGA (*.tga), MNG (*.mng), SVG (*.svg), DOC (*.doc), RTF (*.rtf), TXT (*.txt), WPD (*.wpd), WPS (*.wps), CSV (*.csv), XML (*.xml), XSD (*.xsd), SQL (*.sql), PDF (*.pdf), XLS (*.xls), MDB (*.mdb), PPT (*.ppt), DOCX (*.docx), XLSX (*.xlsx), PPTX (*.pptx), PPSX (*.ppsx), ARTX (*.artx), MP3 (*.mp3), WMA (*.wma), MID (*.mid), MIDI (*.midi), MP4 (*.mp4), MPG (*.mpg), MPEG (*.mpeg), WAV (*.wav), RAM (*.ram), RA (*.ra), AVI (*.avi), MOV (*.mov), FLV (*.flv), M4A (*.m4a), M4V (*.m4v), HTM (*.htm), HTML (*.html), CSS (*.css), SWF (*.swf), JS (*.js), RAR (*.rar), ZIP (*.zip), TAR (*.tar), GZ (*.gz)',
		'fileExt'		 : '*.jpg;*.jpeg;*.jpe;*.jp2;*.jfif;*.gif;*.bmp;*.png;*.psd;*.eps;*.ico;*.tif;*.tiff;*.ai;*.raw;*.tga;*.mng;*.svg;*.doc;*.rtf;*.txt;*.wpd;*.wps;*.csv;*.xml;*.xsd;*.sql;*.pdf;*.xls;*.mdb;*.ppt;*.docx;*.xlsx;*.pptx;*.ppsx;*.artx;*.mp3;*.wma;*.mid;*.midi;*.mp4;*.mpg;*.mpeg;*.wav;*.ram;*.ra;*.avi;*.mov;*.flv;*.m4a;*.m4v;*.htm;*.html;*.css;*.swf;*.js;*.rar;*.zip;*.tar;*.gz',
		'folder'         : '<%=application("uploadpath")%>',
		'multi'          : true,
		onError: function (a, b, c, d) {
         if (d.status == 404)
            alert('Could not find upload script. Use a path relative to: '+'<?= getcwd() ?>');
         else if (d.type === "HTTP")
            alert('error aaa'+d.type+": "+d.status);
         else if (d.type ==="File Size")
            alert(c.name+' '+d.type+' Limit: '+Math.round(d.sizeLimit/1024)+'KB');
         else
            alert('error '+d.type+": "+d.text);
},
		onComplete		 : function(event, queueID, fileObj, response, data) {
     							var path = escape(fileObj.filePath);
								$('#filesUploaded').append('<div class=\'uploadifyQueueItem\'><a href='+path+' target=\'_blank\'>'+fileObj.name+'</a></div>');
								
								location.reload(); 
							}
	});
});
</script>
<%
application("sessionID")	= Session.SessionID
application("uploadpath")	= "damage/" & request("id") & ""

Sub getDamagedItem
	dim intDamageID
	intDamageID = request("id")

    call OpenDataBase()

	set rs = Server.CreateObject("ADODB.recordset")

	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic

	strSQL = "SELECT * FROM yma_damage WHERE damage_id = " & intDamageID

	rs.Open strSQL, conn

    if not DB_RecSetIsEmpty(rs) Then
		session("damage_item") 			= rs("damage_item")
		session("damage_serial_no") 	= rs("damage_serial_no")
		session("damage_type") 			= rs("damage_type")
		session("damage_location")		= rs("damage_location")
		session("damage_connote") 		= rs("damage_connote")
		session("course_damage") 		= rs("course_damage")
		session("sent_excel") 			= rs("sent_excel")
		session("sent_excel_date") 		= rs("sent_excel_date")
		session("lic") 					= rs("lic")
		session("damage_status") 		= rs("status")
		session("damage_date_created") 	= rs("date_created")
		session("damage_created_by") 	= rs("created_by")
		session("damage_date_modified") = rs("date_modified")
		session("damage_modified_by") 	= rs("modified_by")
		session("damage_comments") 		= rs("damage_comments")
    end if

    call CloseDataBase()
end sub

sub updateDamagedItem
	dim strSQL
	dim intDamageID
	intDamageID = request("id")

	dim strItemName
	dim strSerialNo
	dim strDamageType
	dim strDamageLocation
	dim strDamageConnote
	dim strCourseDamage
	dim strSentExcel
	dim strSentExcelDate
	dim strComments

	strItemName 		= trim(Request.Form("txtItemName"))
	strSerialNo 		= trim(Request.Form("txtSerialNo"))
	strDamageType 		= trim(Request.Form("cboDamageType"))
	strDamageLocation 	= trim(Request.Form("txtLocation"))
	strDamageConnote 	= trim(Request.Form("txtConnote"))
	strCourseDamage 	= trim(Request.Form("cboCourseDamage"))
	strSentExcel 		= trim(Request.Form("cboSentExcel"))
	strSentExcelDate 	= trim(Request.Form("txtSentExcelDate"))
	strComments 		= Replace(Request.Form("txtComments"),"'","''")
	
	Call OpenDataBase()

	strSQL = "UPDATE yma_damage SET "
	strSQL = strSQL & "damage_item = '" & Server.HTMLEncode(strItemName) & "',"
	strSQL = strSQL & "damage_serial_no = '" & Server.HTMLEncode(strSerialNo) & "',"
	if Session("UsrLoginRole") = 1 then
		strSQL = strSQL & "lic = '" & trim(Request.Form("txtLIC")) & "',"
	end if
	strSQL = strSQL & "damage_type = '" & strDamageType & "',"
	strSQL = strSQL & "damage_location = '" & strDamageLocation & "',"
	strSQL = strSQL & "damage_connote = '" & strDamageConnote & "',"
	strSQL = strSQL & "course_damage = '" & strCourseDamage & "',"
	strSQL = strSQL & "sent_excel = '" & strSentExcel & "',"
	strSQL = strSQL & "sent_excel_date = CONVERT(datetime,'" & strSentExcelDate & "',103),"
	strSQL = strSQL & "damage_comments = '" & Server.HTMLEncode(strComments) & "',"
	strSQL = strSQL & "date_modified = getdate(),"
	strSQL = strSQL & "modified_by = '" & session("UsrUserName") & "',"
	strSQL = strSQL & "status = '" & trim(Request.Form("cboStatus")) & "' WHERE damage_id = " & intDamageID

	'response.Write strSQL
	on error resume next
	conn.Execute strSQL

	if err <> 0 then
		strMessageText = err.description
	else
		strMessageText = "The Damaged Item has been updated."
	end if

	Call CloseDataBase()

end sub

sub main
	call UTL_validateLogin
	
	dim intID
	intID 	= request("id")
	
	call getDamagedItem
	call getDamageType
	call getCourseDamage
	
	call listComments(intID,warehouseDamageModuleID)
		
	if Request.ServerVariables("REQUEST_METHOD") = "POST" then	
		select case Trim(Request("Action"))
			case "Update"
				call updateDamagedItem
				call getDamagedItem
				call getDamageType
				call getCourseDamage
			case "Comment"
				call addComment(intID,warehouseDamageModuleID)
				call listComments(intID,warehouseDamageModuleID)
		end select
	end if
end sub

call main

dim strMessageText
dim strDamageTypeList
dim strCourseDamageList
dim strCommentsList
%>
</head>
<body>
<table cellspacing="0" cellpadding="0" align="center" class="full_size_table" border="0">
  <!-- #include file="include/header.asp" -->
  <tr>
    <td class="first_content"><table cellpadding="5" cellspacing="0" border="0">
        <tr>
          <td><a href="list_damage.asp"><img src="images/icon_damaged.jpg" border="0" alt="Damage Items" /></a></td>
          <td valign="top"><img src="images/backward_arrow.gif" border="0" /> <a href="list_damage.asp">Back to List</a>
            <h2>Update Warehouse Damage</h2>
            <font color="green"><%= strMessageText %></font></td>
          <td valign="top"><table cellpadding="4" cellspacing="0" class="created_table">
              <tr>
                <td class="created_column_1"><strong>Created:</strong></td>
                <td class="created_column_2"><%= session("damage_created_by") %></td>
                <td class="created_column_3"><%= displayDateFormatted(session("damage_date_created")) %></td>
              </tr>
              <tr>
                <td><strong>Last modified:</strong></td>
                <td><%= session("damage_modified_by") %></td>
                <td><%= displayDateFormatted(session("damage_date_modified")) %></td>
              </tr>
            </table></td>
        </tr>
      </table>
      <table>
        <tr>
          <td width="50%"><form action="" method="post" name="form_update_damage" id="form_update_damage" onsubmit="return validateFormOnSubmit(this)">
              <table cellpadding="5" cellspacing="0" class="item_maintenance_box">
                <tr>
                  <td colspan="2" class="item_maintenance_header">Warehouse Damage</td>
                </tr>
                <tr>
                  <td width="50%">Item name<span class="mandatory">*</span>:<br />
                    <input type="text" id="txtItemName" name="txtItemName" maxlength="20" size="30" value="<%= session("damage_item") %>" /></td>
                  <td width="50%">Serial no<span class="mandatory">*</span>:<br />
                    <input type="text" id="txtSerialNo" name="txtSerialNo" maxlength="20" size="30" value="<%= session("damage_serial_no") %>" /></td>
                </tr>
                <% if Session("UsrLoginRole") = 1 then %>
                <tr>
                  <td colspan="2" bgcolor="#99CCFF">LIC: $
                    <input type="text" id="txtLIC" name="txtLIC" maxlength="10" size="10" value="<%= session("lic") %>" /></td>
                </tr>
                <% end if %>
                <tr>
                  <td>Damage type:<br />
                    <select name="cboDamageType">
                      <%= strDamageTypeList %>
                    </select></td>
                  <td>Cause:<br />
                    <select name="cboCourseDamage">
                      <%= strCourseDamageList %>
                    </select></td>
                </tr>
                <tr>
                  <td>Location:<br />
                    <input type="text" id="txtLocation" name="txtLocation" maxlength="20" size="30" value="<%= session("damage_location") %>" /></td>
                  <td>Con-note:<br />
                    <input type="text" id="txtConnote" name="txtConnote" maxlength="20" size="30" value="<%= session("damage_connote") %>" /></td>
                </tr>
                <tr>
                  <td colspan="2">Sent to Excel:
                    <select name="cboSentExcel">
                      <option <% if session("sent_excel") = "0" then Response.Write " selected" end if%> value="0" rel="none">No</option>
                      <option <% if session("sent_excel") = "1" then Response.Write " selected" end if%> value="1" rel="date">Yes</option>
                    </select></td>
                </tr>
                <tr rel="date">
                  <td colspan="2">Date:
                    <input type="text" id="txtSentExcelDate" name="txtSentExcelDate" maxlength="10" size="10" value="<%= session("sent_excel_date") %>" />
                    <em>DD/MM/YYYY</em></td>
                </tr>
                <tr>
                  <td colspan="2">Comments:<br />
                    <textarea name="txtComments" id="txtComments" cols="55" rows="3"><%= session("damage_comments") %></textarea></td>
                </tr>
                <tr class="status_row">
                  <td colspan="2">Status:
                    <select name="cboStatus">
                      <option <% if session("damage_status") = "1" then Response.Write " selected" end if%> value="1">Open</option>
                      <option <% if session("damage_status") = "0" then Response.Write " selected" end if%> value="0">Completed</option>
                    </select></td>
                </tr>
                <tr>
                  <td colspan="2"><input type="hidden" name="Action" />
                    <input type="submit" value="Update Warehouse Damage" /></td>
                </tr>
              </table>
            </form></td>
          <td width="50%" valign="top" style="padding-left:15px;"><table cellpadding="5" cellspacing="0" class="serial_no_box">
              <tr>
                <td class="item_maintenance_header">Documents <img src="images/icon_new.gif" border="0" align="top" /></td>
              </tr>
              <tr>
                <td><form id="formIDdoc" name="formIDdoc" class="form" method="post">             
                <p>Please select the files first, then click the "Upload" button.</p>       
                    <p>
                      <input class="text-input" name="uploadify" id="uploadify" type="file" size="20" />
                    </p>
                    <h3 align="right"><a href="javascript:Send_document()"><img src="images/btn_upload.gif" border="0" align="top" /></a></h3>
                    <div id="filesUploaded"></div>
                  </form>
                  <p><% ListFolderContents(Server.MapPath("damage/" & request("id") & "")) %></p></td>
              </tr>
            </table></td>
        </tr>
      </table>
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
</table>
</body>
</html>