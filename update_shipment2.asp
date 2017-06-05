<% response.cookies("current_URL_cookie_logistics") = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?" & Request.Querystring %>
<!--#include file="include/connection_it.asp " -->
<!--#include file="class/clsComment.asp " -->
<% strSection = "shipment" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Update Shipment</title>
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="include/generic_form_validations.js"></script>
<script language="JavaScript" type="text/javascript">
function validateFormOnSubmit(theForm) {
	var reason = "";
	var blnSubmit = true;

	reason += validateEmptyField(theForm.txtContainerNo,"Container No");
	reason += validateSpecialCharacters(theForm.txtContainerNo,"Container No");
	reason += validateEmptyField(theForm.txtSupplierInvoiceNo,"Supplier Invoice No");
	reason += validateSpecialCharacters(theForm.txtSupplierInvoiceNo,"Supplier Invoice No");
	reason += validateSpecialCharacters(theForm.txtEFT,"EFT");
	reason += validateSpecialCharacters(theForm.txtCommodity,"Commodity");
	reason += validateSpecialCharacters(theForm.txtPortOrigin,"Port of Origin");
	reason += validateSpecialCharacters(theForm.txtVesselName,"Vessel Name");
	reason += validateSpecialCharacters(theForm.txtVoyage,"Voyage");
	reason += validateEmptyField(theForm.txtCartons,"No of Cartons");
	reason += validateEmptyField(theForm.cboTEU,"TEU");
	reason += validateEmptyField(theForm.txtInvoice1,"Invoice 1");	

  	if (reason != "") {
    	alert("Some fields need correction:\n" + reason);

		blnSubmit = false;
		return false;
  	}

	if (blnSubmit == true){
        theForm.Action.value = 'Update';
  		theForm.submit();

		return true;
    }
}

function copyShipment(theForm) {
	if (confirm ("Please click OK to copy this shipment.")){
		theForm.Action.value = 'Copy';
		return true;
    } else {
		return false;
	}
}

function submitCertificateOrigin(theForm) {
	var reason = "";
	var blnSubmit = true;
	
	reason += validateCheckBox(theForm.chkCertificateOrigin,"Certificate of Origin");
	
	if (reason != "") {
    	alert("Some fields need correction:\n" + reason);

		blnSubmit = false;
		return false;
  	}
	
	if (blnSubmit == true){
        theForm.Action.value = 'Update Certificate Origin';
  		theForm.submit();

		return true;
    }
}

function submitRefundApplication(theForm) {
	var reason = "";
	var blnSubmit = true;
	
	reason += validateCheckBox(theForm.chkRefundApplication,"Refund Application");
	
	if (reason != "") {
    	alert("Some fields need correction:\n" + reason);

		blnSubmit = false;
		return false;
  	}

	if (blnSubmit == true){
        theForm.Action.value = 'Update Refund Application';
  		theForm.submit();

		return true;
    }
}

function submitImportDeclaration(theForm) {
	var reason = "";
	var blnSubmit = true;
	
	reason += validateCheckBox(theForm.chkImportDeclaration,"Import Declaration");
	
	if (reason != "") {
    	alert("Some fields need correction:\n" + reason);

		blnSubmit = false;
		return false;
  	}

	if (blnSubmit == true){
        theForm.Action.value = 'Update Import Declaration';
  		theForm.submit();

		return true;
    }
}

function submitRefund(theForm) {
	var reason = "";
	var blnSubmit = true;
	
	reason += validateCheckBox(theForm.chkRefund,"Refund");
	
	if (reason != "") {
    	alert("Some fields need correction:\n" + reason);

		blnSubmit = false;
		return false;
  	}

	if (blnSubmit == true){
        theForm.Action.value = 'Update Refund';
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
	
	if (blnSubmit == true){
		theForm.Action.value = 'Comment';
		
		return true;		
    }
}
</script>
<%
sub ListFolderContents(path)
	dim fs, folder, file, item, url

    set fs = CreateObject("Scripting.FileSystemObject")
    set folder = fs.GetFolder(path)

    'Display the target folder and info.

    Response.Write("Summary: " _
       & folder.Files.Count & " files, ")
    if folder.SubFolders.Count > 0 then
       Response.Write(folder.SubFolders.Count & " directories, ")
    end if
    Response.Write(Round(folder.Size / 1024) & " KB total." _
       & vbCrLf)

    Response.Write("<ul>" & vbCrLf)

    'Display a list of sub folders.

    for each item in folder.SubFolders
    	ListFolderContents(item.Path)
    next

    'Display a list of files.

    for each item in folder.Files
       	url = MapURL(item.path)
    	Response.Write("<li><a href=""" & url & """>" & item.Name & "</a>" & vbCrLf)
    next

	Response.Write("</ul>" & vbCrLf)
    Response.Write("</li>" & vbCrLf)
end sub

function MapURL(path)
    dim rootPath, url

    'Convert a physical file path to a URL for hypertext links.

    rootPath = Server.MapPath("/")
    url = Right(path, Len(path) - Len(rootPath))
    MapURL = Replace(url, "\", "/")
end function

'-----------------------------------------------
' RETRIEVE SHIPMENT RECORDS
'-----------------------------------------------

Sub getShipment

	dim intID
	intID = request("id")

    call OpenDataBase()

	set rs = Server.CreateObject("ADODB.recordset")

	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic

	strSQL = "SELECT * FROM yma_shipment WHERE shipment_id = " & intID

	rs.Open strSQL, conn

	'Response.Write strSQL

    if not DB_RecSetIsEmpty(rs) Then
		session("container_no") 			= rs("container_no")
		session("supplier_invoice_no") 		= rs("supplier_invoice_no")
		session("department") 				= rs("department")
		session("supplier_contact") 		= rs("supplier_contact")
		session("custom_cleared") 			= rs("custom_cleared")
		session("fumigation") 				= rs("fumigation")
		session("air_freight") 				= rs("air_freight")
		session("priority") 				= rs("priority")
		session("edo") 						= rs("edo")
		session("FF") 						= rs("FF")
		session("EFT") 						= rs("EFT")
		session("all_documents") 			= rs("all_documents")
		session("bill_lading") 				= rs("bill_lading")
		session("bill_lading_date") 		= rs("bill_lading_date")
		session("commercial_invoice") 		= rs("commercial_invoice")
		session("commercial_invoice_date") 	= rs("commercial_invoice_date")
		session("packing_list") 			= rs("packing_list")
		session("packing_list_date") 		= rs("packing_list_date")
		session("packing_declaration") 		= rs("packing_declaration")
		session("packing_declaration_date") = rs("packing_declaration_date")
		session("treatment_certificate") 	= rs("treatment_certificate")
		session("treatment_certificate_date") = rs("treatment_certificate_date")
		session("certificate_origin") 		= rs("certificate_origin")
		session("certificate_origin_date") 	= rs("certificate_origin_date")
		session("comments") 				= rs("comments")
		session("commodity") 				= rs("commodity")
		session("port_origin") 				= rs("port_origin")
		session("country_origin") 			= rs("country_origin")
		session("vessel_name") 				= rs("vessel_name")
		session("voyage") 					= rs("voyage")
		session("warehouse") 				= rs("warehouse")
		session("cartons") 					= rs("cartons")
		session("date_shipment") 			= rs("date_shipment")
		session("eta_discharged") 			= rs("eta_discharged")
		session("melb_eta_time") 			= rs("melb_eta_time")
		session("eta_availability") 		= rs("eta_availability")
		session("eta_unpacked") 			= rs("eta_unpacked")
		session("teu") 						= rs("teu")
		session("invoice1") 				= rs("invoice1")
		session("invoice2") 				= rs("invoice2")
		session("invoice3") 				= rs("invoice3")
		session("invoice4") 				= rs("invoice4")
		session("invoice5") 				= rs("invoice5")
		session("fta_certificate_origin") 		= rs("fta_certificate_origin")
		session("fta_certificate_origin_date")	= rs("fta_certificate_origin_date")
		session("fta_refund_application") 		= rs("fta_refund_application")
		session("fta_refund_application_date") 	= rs("fta_refund_application_date")
		session("fta_import_declaration") 		= rs("fta_import_declaration")
		session("fta_import_declaration_date") 	= rs("fta_import_declaration_date")
		session("fta_refund") 					= rs("fta_refund")
		session("fta_refund_date") 				= rs("fta_refund_date")
		session("fta_status") 					= rs("fta_status")
		session("status") 				= rs("status")
		session("modified_date") 		= rs("modified_date")
		session("modified_by") 			= rs("modified_by")
		session("date_created") 		= rs("date_created")
		session("created_by") 			= rs("created_by")
    end if

    call CloseDataBase()

end sub

'-----------------------------------------------
' UPDATE SHIPMENT
'-----------------------------------------------

sub updateShipment
	dim strSQL
	dim intID
	intID = request("id")
	
	dim strContainer
	dim strSupplierInvoiceNo
	dim strDepartment
	dim strSupplierContact
	dim strCustomCleared
	dim strFumigation
	dim intAirFreight
	dim intPriority
	dim intEDO
	dim strFF
	dim strEFT
	dim strAllDocuments
	dim intBillLading
	dim intCommercialInvoice
	dim intPackingList
	dim intPackingDeclaration
	dim intTreatmentCertificate
	dim intCertificateOrigin
	dim strComments
	dim strCommodity
	dim strPortOrigin
	dim strCountryOrigin
	dim strVesselName
	dim strVoyage
	dim strWarehouse
	dim intCartons
	dim strDateShipment
	dim strEtaDischarged
	dim strEtaAvailability
	dim strEtaUnpacked
	dim strContainerEtaTime
	dim strTEU
	dim strInvoice1
	dim strInvoice2
	dim strInvoice3
	dim strInvoice4
	dim strInvoice5
	dim intStatus
	
	strContainer 			= trim(Request.Form("txtContainerNo"))
	strSupplierInvoiceNo 	= trim(Request.Form("txtSupplierInvoiceNo"))	
	strDepartment 			= trim(Request.Form("cboDepartment"))
	strSupplierContact 		= trim(Request.Form("cboSupplierContact"))
	strCustomCleared 		= trim(Request.Form("cboCustomCleared"))	
	strFumigation 			= trim(Request.Form("cboFumigation"))	
	intAirFreight 			= trim(Request.Form("chkAirFreight"))
	intPriority 			= trim(Request.Form("chkPriority"))
	intEDO 					= trim(Request.Form("chkEDO"))
	strFF 					= trim(Request.Form("cboFF"))
	strEFT 					= trim(Request.Form("txtEFT"))
	strAllDocuments 		= trim(Request.Form("cboAllDocuments"))
	intBillLading 			= trim(Request.Form("chkBillLading"))
	intCommercialInvoice 	= trim(Request.Form("chkCommercialInvoice"))
	intPackingList 			= trim(Request.Form("chkPackingList"))
	intPackingDeclaration 	= trim(Request.Form("chkPackingDeclaration"))
	intTreatmentCertificate = trim(Request.Form("chkTreatmentCertificate"))
	intCertificateOrigin 	= trim(Request.Form("chkCertificateOrigin"))
	strComments 			= Replace(Request.Form("txtComments"),"'","''")
	strCommodity 			= trim(Request.Form("txtCommodity"))	
	strPortOrigin 			= trim(Request.Form("txtPortOrigin"))	
	strCountryOrigin 		= trim(Request.Form("cboCountryOrigin"))	
	strVesselName 			= trim(Request.Form("txtVesselName"))	
	strVoyage 				= trim(Request.Form("txtVoyage"))	
	strWarehouse 			= trim(Request.Form("cboWarehouse"))
	intCartons 				= trim(Request.Form("txtCartons"))
	strDateShipment 		= trim(Request.Form("txtDateShipment"))	
	strEtaDischarged 		= trim(Request.Form("txtEtaDischarged"))	
	strEtaAvailability 		= trim(Request.Form("txtEtaAvailability"))	
	strEtaUnpacked 			= trim(Request.Form("txtEtaUnpacked"))	
	strContainerEtaTime 	= trim(Request.Form("txtContainerEtaTime"))
	strTEU 					= trim(Request.Form("cboTEU"))	
	strInvoice1 			= trim(Request.Form("txtInvoice1"))
	strInvoice2 			= trim(Request.Form("txtInvoice2"))
	strInvoice3 			= trim(Request.Form("txtInvoice3"))
	strInvoice4 			= trim(Request.Form("txtInvoice4"))
	strInvoice5 			= trim(Request.Form("txtInvoice5"))
	intStatus				= trim(Request.Form("cboStatus"))
	
	Call OpenDataBase()

	strSQL = "UPDATE yma_shipment SET "
	strSQL = strSQL & "container_no = '" & strContainer & "',"
	strSQL = strSQL & "supplier_invoice_no = '" & strSupplierInvoiceNo & "',"
	strSQL = strSQL & "department = '" & strDepartment & "',"
	strSQL = strSQL & "supplier_contact = '" & strSupplierContact & "',"
	strSQL = strSQL & "custom_cleared = '" & strCustomCleared & "',"
	strSQL = strSQL & "fumigation = '" & strFumigation & "',"
	strSQL = strSQL & "air_freight = '" & intAirFreight & "',"
	strSQL = strSQL & "priority = '" & intPriority & "',"
	strSQL = strSQL & "edo = '" & intEDO & "',"
	strSQL = strSQL & "FF = '" & strFF & "',"
	strSQL = strSQL & "EFT = '" & strEFT & "',"
	strSQL = strSQL & "all_documents = '" & strAllDocuments & "',"
	'strSQL = strSQL & "comments = '" & Server.HTMLEncode(strComments) & "',"
	strSQL = strSQL & "commodity = '" & strCommodity & "',"
	strSQL = strSQL & "port_origin = '" & strPortOrigin & "',"
	strSQL = strSQL & "country_origin = '" & strCountryOrigin & "',"
	strSQL = strSQL & "vessel_name = '" & strVesselName & "',"
	strSQL = strSQL & "voyage = '" & strVoyage & "',"
	strSQL = strSQL & "warehouse = '" & strWarehouse & "',"
	strSQL = strSQL & "cartons = '" & intCartons & "',"
	strSQL = strSQL & "date_shipment = CONVERT(datetime,'" & strDateShipment & "',103),"
	strSQL = strSQL & "eta_discharged = CONVERT(datetime,'" & trim(Request.Form("txtMelbETA")) & "',103),"
	strSQL = strSQL & "eta_availability = CONVERT(datetime,'" &	trim(Request.Form("txtContainerETA")) & "',103),"
	strSQL = strSQL & "eta_unpacked = CONVERT(datetime,'" & trim(Request.Form("txtUnpackETA")) & "',103),"
	strSQL = strSQL & "melb_eta_time = '" & strContainerEtaTime & "',"
	strSQL = strSQL & "bill_lading = '" & intBillLading & "',"
	strSQL = strSQL & "commercial_invoice = '" & intCommercialInvoice & "',"
	strSQL = strSQL & "packing_list = '" & intPackingList & "',"
	strSQL = strSQL & "packing_declaration = '" & intPackingDeclaration & "',"
	strSQL = strSQL & "treatment_certificate = '" & intTreatmentCertificate & "',"
	strSQL = strSQL & "certificate_origin = '" & intCertificateOrigin & "',"
	strSQL = strSQL & "teu = '" & strTEU & "',"
	strSQL = strSQL & "invoice1 = '" & strInvoice1 & "',"
	strSQL = strSQL & "invoice2 = '" & strInvoice2 & "',"
	strSQL = strSQL & "invoice3 = '" & strInvoice3 & "',"
	strSQL = strSQL & "invoice4 = '" & strInvoice4 & "',"
	strSQL = strSQL & "invoice5 = '" & strInvoice5 & "',"
	strSQL = strSQL & "modified_date = getdate(),"
	strSQL = strSQL & "modified_by = '" & session("UsrUserName") & "',"
	strSQL = strSQL & "status = '" & intStatus & "' WHERE shipment_id = " & intID

	'response.Write strSQL
	on error resume next
	conn.Execute strSQL

	if err <> 0 then
		strMessageText = err.description
	else
		strMessageText = "Shipment has been updated."
	end if

	Call CloseDataBase()
end sub

'-----------------------------------------------
' COPY THIS SHIPMENT RECORD
'-----------------------------------------------

sub copyShipment
	dim strSQL
	
	call OpenDataBase()
		
strSQL = "INSERT INTO yma_shipment (container_no, supplier_invoice_no, department, supplier_contact, custom_cleared, fumigation, air_freight, priority, FF, EFT, all_documents, bill_lading,  commercial_invoice, packing_list, packing_declaration, treatment_certificate, certificate_origin, comments, commodity, port_origin, country_origin, vessel_name, voyage, warehouse, cartons, date_shipment, eta_discharged, eta_availability, eta_unpacked, melb_eta_time, teu, date_created, created_by, invoice1, invoice2, invoice3, invoice4, invoice5) VALUES ( "
	strSQL = strSQL & "'" & session("container_no")  & " COPY',"	
	strSQL = strSQL & "'" & session("supplier_invoice_no") & " COPY',"	
	strSQL = strSQL & "'" & session("department") & "',"
	strSQL = strSQL & "'" & session("supplier_contact") & "',"
	strSQL = strSQL & "'" & session("custom_cleared") & "',"	
	strSQL = strSQL & "'" & session("fumigation") & "',"	
	strSQL = strSQL & "'" & session("air_freight") & "',"
	strSQL = strSQL & "'" & session("priority") & "',"
	strSQL = strSQL & "'" & session("FF") & "',"
	strSQL = strSQL & "'" & session("EFT")  & "',"
	strSQL = strSQL & "'" & session("all_documents") & "',"
	strSQL = strSQL & "'" & session("bill_lading") & "',"
	strSQL = strSQL & "'" & session("commercial_invoice") & "',"
	strSQL = strSQL & "'" & session("packing_list") & "',"
	strSQL = strSQL & "'" & session("packing_declaration") & "',"
	strSQL = strSQL & "'" & session("treatment_certificate") & "',"
	strSQL = strSQL & "'" & session("certificate_origin") & "',"
	strSQL = strSQL & "'" & session("comments") & "',"
	strSQL = strSQL & "'" & session("commodity") & "',"	
	strSQL = strSQL & "'" & session("port_origin") & "',"	
	strSQL = strSQL & "'" & session("country_origin") & "',"	
	strSQL = strSQL & "'" & session("vessel_name") & "',"	
	strSQL = strSQL & "'" & session("voyage") & "',"	
	strSQL = strSQL & "'" & session("warehouse") & "',"
	strSQL = strSQL & "'" & session("cartons") & "',"	
	strSQL = strSQL & " CONVERT(datetime,'" & session("date_shipment") & "',103),"
	strSQL = strSQL & " CONVERT(datetime,'" & session("eta_discharged")  & "',103),"
	strSQL = strSQL & " CONVERT(datetime,'" & session("eta_availability") & "',103),"	
	strSQL = strSQL & " CONVERT(datetime,'" & session("eta_unpacked") & "',103),"
	strSQL = strSQL & "'" & session("melb_eta_time") & "',"
	strSQL = strSQL & "'" & session("teu") & "',getdate(),"	
	strSQL = strSQL & "'" & session("UsrUserName") & "',"
	strSQL = strSQL & "'" & session("invoice1") & "',"
	strSQL = strSQL & "'" & session("invoice2") & "',"
	strSQL = strSQL & "'" & session("invoice3") & "',"
	strSQL = strSQL & "'" & session("invoice4") & "',"
	strSQL = strSQL & "'" & session("invoice5") & "')"
	
	'response.Write strSQL 
	on error resume next
	conn.Execute strSQL
		
	if err <> 0 then
		strMessageText = err.description
	else 	
		Response.Redirect("thank-you_shipment.asp")
	end if 

	call CloseDataBase()
end sub

'-----------------------------------------------
' UPDATE FTA: CERTIFICATE OF ORIGIN
'-----------------------------------------------

sub updateCertificateOrigin
	dim strSQL
	dim intID
	intID = request("id")

	Call OpenDataBase()

	strSQL = "UPDATE yma_shipment SET "
	strSQL = strSQL & "fta_certificate_origin = '" & trim(Request.Form("chkCertificateOrigin")) & "',"
	strSQL = strSQL & "fta_certificate_origin_date =  getdate(), fta_status = 1 WHERE shipment_id = " & intID

	'response.Write strSQL
	on error resume next
	conn.Execute strSQL

	On error Goto 0

	if err <> 0 then
		strMessageText = err.description
	else
		strMessageText = "FTA: Certificate of Origin has been updated."
	end if

	Call CloseDataBase()
end sub

'-----------------------------------------------
' UPDATE FTA: REFUND APPLICATION
'-----------------------------------------------

sub updateRefundApplication
	dim strSQL
	dim intID
	intID = request("id")

	Call OpenDataBase()

	strSQL = "UPDATE yma_shipment SET "
	strSQL = strSQL & "fta_refund_application = '" & trim(Request.Form("chkRefundApplication")) & "',"
	strSQL = strSQL & "fta_refund_application_date =  getdate(), fta_status = 2 WHERE shipment_id = " & intID

	'response.Write strSQL
	on error resume next
	conn.Execute strSQL

	On error Goto 0

	if err <> 0 then
		strMessageText = err.description
	else
		strMessageText = "FTA: Refund Application has been updated."
	end if

	Call CloseDataBase()
end sub

'-----------------------------------------------
' UPDATE FTA: IMPORT DECLARATION
'-----------------------------------------------

sub updateImportDeclaration
	dim strSQL
	dim intID
	intID = request("id")

	Call OpenDataBase()

	strSQL = "UPDATE yma_shipment SET "
	strSQL = strSQL & "fta_import_declaration = '" & trim(Request.Form("chkImportDeclaration")) & "',"
	strSQL = strSQL & "fta_import_declaration_date =  getdate(), fta_status = 3 WHERE shipment_id = " & intID

	'response.Write strSQL
	on error resume next
	conn.Execute strSQL

	On error Goto 0

	if err <> 0 then
		strMessageText = err.description
	else
		strMessageText = "FTA: Import Declaration has been updated."
	end if

	Call CloseDataBase()
end sub

'-----------------------------------------------
' UPDATE FTA: REFUND
'-----------------------------------------------

sub updateRefund
	dim strSQL
	dim intID
	intID = request("id")

	Call OpenDataBase()

	strSQL = "UPDATE yma_shipment SET "
	strSQL = strSQL & "fta_refund = '" & trim(Request.Form("chkRefund")) & "',"
	strSQL = strSQL & "fta_refund_date =  getdate(), fta_status = 4 WHERE shipment_id = " & intID

	'response.Write strSQL
	on error resume next
	conn.Execute strSQL

	On error Goto 0

	if err <> 0 then
		strMessageText = err.description
	else
		strMessageText = "FTA: Refund has been updated."
	end if

	Call CloseDataBase()
end sub

sub main
	call UTL_validateLogin
	
	dim intID
	intID 	= request("id")
	
	call getShipment
	call listComments(intID,shipmentModuleID)
	
	select case Trim(Request("ref"))
		case "open"
			strBackLink = "list_shipment.asp"
		case "past"	
			strBackLink = "list_past-shipment.asp"
		case else
			strBackLink = "list_shipment.asp"
	end select
	
	if Request.ServerVariables("REQUEST_METHOD") = "POST" then
		select case Trim(Request("Action"))
			case "Update"
				call updateShipment
				call getShipment			
			case "Copy"
				call copyShipment
			case "Update Certificate Origin"
				call updateCertificateOrigin
				call getShipment
			case "Update Refund Application"
				call updateRefundApplication
				call getShipment
			case "Update Import Declaration"
				call updateImportDeclaration
				call getShipment
			case "Update Refund"
				call updateRefund
				call getShipment
			case "Comment"
				call addComment(intID,shipmentModuleID)
				call listComments(intID,shipmentModuleID)
		end select
	end if
	
end sub

dim strBackLink
dim strMessageText
dim strCommentsList

call main
%>
</head>
<body>
<table cellspacing="0" cellpadding="0" align="center" class="full_size_table" border="0">
  <!-- #include file="include/header.asp" -->
  <tr>
    <td class="first_content"><table cellpadding="5" cellspacing="0" border="0">
        <tr>
          <td valign="top"><a href="list_shipment.asp"><img src="images/icon_shipment.jpg" border="0" alt="Shipment" /></a></td>
          <td valign="top"><img src="images/backward_arrow.gif" border="0" /> <a href="<%= strBackLink %>">Back to List</a>
            <h2>Update Shipment</h2>
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
                <td><%= displayDateFormatted(session("modified_date")) %></td>
              </tr>
            </table>
            <form action="" method="post" name="form_copy_shipment" id="form_copy_shipment" onsubmit="return copyShipment(this)">
              <p>
                <input type="hidden" name="Action" />
                <input type="submit" value="Copy this shipment" />
                <img src="images/icon_new.gif" border="0" align="top" /></p>
            </form></td>
        </tr>
      </table>
      <table cellpadding="0" cellspacing="0" border="0">
        <tr>
          <td><form action="" method="post" name="form_update_shipment" id="form_update_shipment" onsubmit="return validateFormOnSubmit(this)">
              <table border="0" cellpadding="5" cellspacing="0" width="1000">
                <tr>
                  <td width="50%" valign="top"><table cellpadding="5" cellspacing="0" class="item_maintenance_box">
                      <tr>
                        <td colspan="3" class="item_maintenance_header">Container Info</td>
                      </tr>
                      <tr>
                        <td>Container no<span class="mandatory">*</span>:<br />
                          <input type="text" id="txtContainerNo" name="txtContainerNo" maxlength="20" size="20" value="<%= Server.HTMLEncode(session("container_no")) %>" />
                          <em>(TBA if blank)</em></td>
                        <td colspan="2">Supplier invoice no<span class="mandatory">*</span>:<br />
                          <input type="text" id="txtSupplierInvoiceNo" name="txtSupplierInvoiceNo" maxlength="20" size="20" value="<%= Server.HTMLEncode(session("supplier_invoice_no")) %>" /></td>
                      </tr>
                      <tr>
                        <td>Department<span class="mandatory">*</span>:<br />
                          <select name="cboDepartment">
                            <option <% if session("department") = "AV" then Response.Write " selected" end if%> value="AV">AV</option>
                            <option <% if session("department") = "MPD" then Response.Write " selected" end if%> value="MPD">MPD</option>
                            <option <% if session("department") = "AV-MPD" then Response.Write " selected" end if%> value="AV-MPD">AV &amp; MPD</option>
                            <option <% if session("department") = "Other" then Response.Write " selected" end if%> value="Other">Other</option>
                          </select></td>
                        <td colspan="2">Supplier contact:<br />
                          <select name="cboSupplierContact">
                            <option <% if session("supplier_contact") = "" then Response.Write " selected" end if%> value="">...</option>
                            <option <% if session("supplier_contact") = "distorder@steinberg.de" then Response.Write " selected" end if%> value="distorder@steinberg.de">Steinberg (Annegret)</option>
                            <option <% if session("supplier_contact") = "jwigger@paiste.com" then Response.Write " selected" end if%> value="jwigger@paiste.com">Paiste (Jasmin)</option>
                            <option <% if session("supplier_contact") = "mark@belcat.com" then Response.Write " selected" end if%> value="mark@belcat.com">Belcat (Mark)</option>
                            <option <% if session("supplier_contact") = "takaoka@korg.co.jp" then Response.Write " selected" end if%> value="takaoka@korg.co.jp">Korg (Shigeru)</option>
                            <option <% if session("supplier_contact") = "fujiwara@korg.co.jp" then Response.Write " selected" end if%> value="fujiwara@korg.co.jp">Korg (Shungo)</option>
                            <option <% if session("supplier_contact") = "mamii@korg.jp" then Response.Write " selected" end if%> value="mamii@korg.jp">Korg (Naoko)</option>
                            <option <% if session("supplier_contact") = "khsms1@khs-musix.com" then Response.Write " selected" end if%> value="khsms1@khs-musix.com">KHS (Iris)</option>
                            <option <% if session("supplier_contact") = "msorensen@yamaha.com" then Response.Write " selected" end if%> value="msorensen@yamaha.com">YAMUSA (Mellonie)</option>
                            <option <% if session("supplier_contact") = "ismerdel@yamaha.com" then Response.Write " selected" end if%> value="ismerdel@yamaha.com">Everbright (Irene)</option>
                          </select></td>
                      </tr>
                      <tr>
                        <td><input type="checkbox" name="chkAirFreight" id="chkAirFreight" value="1" <% if session("air_freight") = "1" then Response.Write " checked" end if%> />
                          Air freight <img src="images/airplane.gif" alt="" border="0" /></td>
                        <td colspan="2"><input type="checkbox" name="chkPriority" id="chkPriority" value="1" <% if session("priority") = "1" then Response.Write " checked" end if%> />
                          Priority <img src="images/icon_priority.gif" alt="" border="0" /></td>
                      </tr>
                      <tr>
                        <td width="50%">Custom cleared:<br />
                          <select name="cboCustomCleared">
                            <option <% if session("custom_cleared") = "-" then Response.Write " selected" end if%> value="-">...</option>
                            <option <% if session("custom_cleared") = "Y" then Response.Write " selected" end if%> value="Y">Yes</option>
                            <option <% if session("custom_cleared") = "N" then Response.Write " selected" end if%> value="N">No</option>
                            <option <% if session("custom_cleared") = "Border-hold" then Response.Write " selected" end if%> value="Border-hold" style="background-color:#F00; color:#FFF">Border Hold</option>
                            <option <% if session("custom_cleared") = "X-ray" then Response.Write " selected" end if%> value="X-ray">X-ray</option>
                          </select></td>
                        <td width="20%">Fumigation:<br />
                          <select name="cboFumigation">
                            <option <% if session("fumigation") = "-" then Response.Write " selected" end if%> value="-">...</option>
                            <option <% if session("fumigation") = "Y" then Response.Write " selected" end if%> value="Y">Yes</option>
                            <option <% if session("fumigation") = "N" then Response.Write " selected" end if%> value="N">No</option>
                            <option <% if session("fumigation") = "AQIS inspect" then Response.Write " selected" end if%> value="AQIS inspect">AQIS inspect</option>
                          </select></td>
                        <td width="30%">Freight fwd:<br />
                          <select name="cboFF">
                            <option <% if session("FF") = "-" then Response.Write " selected" end if%> value="-">...</option>
                            <option <% if session("FF") = "M" then Response.Write " selected" end if%> value="M">M</option>
                            <option <% if session("FF") = "K" then Response.Write " selected" end if%> value="K">K</option>
                            <option <% if session("FF") = "Y" then Response.Write " selected" end if%> value="Y">Y</option>
                          </select></td>
                      </tr>
                      <tr>
                        <td>EFT:<br />
                          <input type="text" id="txtEFT" name="txtEFT" maxlength="5" size="6" value="<%= session("EFT") %>" /></td>
                        <td colspan="2">All documents:<br />
                          <select name="cboAllDocuments">
                            <option <% if session("all_documents") = "N" then Response.Write " selected" end if%> value="N">No</option>
                            <option <% if session("all_documents") = "Y" then Response.Write " selected" end if%> value="Y">Yes</option>
                            <option <% if session("all_documents") = "Part" then Response.Write " selected" end if%> value="Part" style="background-color:#F00; color:#FFF">Part</option>
                          </select></td>
                      </tr>
                      <tr>
                        <td>Commodity:<br />
                          <input type="text" id="txtCommodity" name="txtCommodity" maxlength="50" size="30" value="<%= Server.HTMLEncode(session("commodity")) %>" /></td>
                        <td colspan="2">TEU<span class="mandatory">*</span>:<br />
                          <select name="cboTEU">
                            <option <% if session("teu") = "" then Response.Write " selected" end if%> value="">...</option>
                            <option <% if session("teu") = "20" then Response.Write " selected" end if%> value="20">20</option>
                            <option <% if session("teu") = "40" then Response.Write " selected" end if%> value="40">40</option>
                            <option <% if session("teu") = "HC" then Response.Write " selected" end if%> value="HC">HC</option>
                            <option <% if session("teu") = "LCL" then Response.Write " selected" end if%> value="LCL">LCL</option>
                            <option <% if session("teu") = "Air Freight" then Response.Write " selected" end if%> value="Air Freight">Air Freight</option>
                          </select></td>
                      </tr>
                      <tr>
                        <td>Port of origin:<br />
                          <input type="text" id="txtPortOrigin" name="txtPortOrigin" maxlength="50" size="20" value="<%= Server.HTMLEncode(session("port_origin")) %>" /></td>
                        <td colspan="2">Country of origin:<br />
                          <select name="cboCountryOrigin">
                            <option <% if session("country_origin") = "NA" then Response.Write " selected" end if%> value="NA">...</option>
                            <option <% if session("country_origin") = "China" then Response.Write " selected" end if%> value="China">China</option>
                            <option <% if session("country_origin") = "England" then Response.Write " selected" end if%> value="England">England</option>
                            <option <% if session("country_origin") = "Germany" then Response.Write " selected" end if%> value="Germany">Germany</option>
                            <option <% if session("country_origin") = "Hongkong" then Response.Write " selected" end if%> value="Hongkong">Hongkong</option>
                            <option <% if session("country_origin") = "Indonesia" then Response.Write " selected" end if%> value="Indonesia">Indonesia</option>
                            <option <% if session("country_origin") = "Japan" then Response.Write " selected" end if%> value="Japan">Japan</option>
                            <option <% if session("country_origin") = "Malaysia" then Response.Write " selected" end if%> value="Malaysia">Malaysia</option>
                            <option <% if session("country_origin") = "NZ" then Response.Write " selected" end if%> value="NZ">NZ</option>
                            <option <% if session("country_origin") = "Singapore" then Response.Write " selected" end if%> value="Singapore">Singapore</option>
                            <option <% if session("country_origin") = "USA" then Response.Write " selected" end if%> value="USA">USA</option>
                            <option <% if session("country_origin") = "Vietnam" then Response.Write " selected" end if%> value="Vietnam">Vietnam</option>
                            <option <% if session("country_origin") = "Other" then Response.Write " selected" end if%> value="Other">Other</option>
                          </select></td>
                      </tr>
                      <tr>
                        <td>Vessel:<br />
                          <input type="text" id="txtVesselName" name="txtVesselName" maxlength="50" size="40" value="<%= Server.HTMLEncode(session("vessel_name")) %>" /></td>
                        <td colspan="2">Voyage:<br />
                          <input type="text" id="txtVoyage" name="txtVoyage" maxlength="8" size="8" value="<%= Server.HTMLEncode(session("voyage")) %>" /></td>
                      </tr>
                      <tr>
                        <td>Warehouse:<br />
                          <select name="cboWarehouse">
                            <option <% if session("warehouse") = "" then Response.Write " selected" end if%> value="">...</option>
                            <option <% if session("warehouse") = "TT" then Response.Write " selected" end if%> value="TT">TT</option>
                            <option <% if session("warehouse") = "EXL" then Response.Write " selected" end if%> value="EXL">EXL</option>
                            <option <% if session("warehouse") = "YMA" then Response.Write " selected" end if%> value="YMA">YMA Head Office</option>
                          </select></td>
                        <td colspan="2">No of cartons<span class="mandatory">*</span>:<br />
                          <input type="text" id="txtCartons" name="txtCartons" maxlength="5" size="5" value="<%= session("cartons") %>" /></td>
                      </tr>
                      <tr>
                        <td colspan="3"><input type="checkbox" name="chkEDO" id="chkEDO" value="1" <% if session("edo") = "1" then Response.Write " checked" end if%> /> EDO <img src="images/icon_new.gif" border="0" align="top" /></td>
                      </tr>
                      <tr>
                        <td colspan="3">Additional Info:</td>
                      </tr>
                      <tr>
                        <td colspan="3" class="comment_column">&nbsp;<em><%= session("comments") %></em></td>
                      </tr>
                      <tr class="status_row">
                        <td colspan="3">Status:
                          <select name="cboStatus">
                            <option <% if session("status") = "1" then Response.Write " selected" end if%> value="1">Open</option>
                            <option <% if session("status") = "0" then Response.Write " selected" end if%> value="0" style="background-color:#0F0">Completed</option>
                          </select></td>
                      </tr>
                    </table></td>
                  <td width="50%" valign="top"><table cellpadding="5" cellspacing="0" class="item_maintenance_box">
                      <tr>
                        <td colspan="2" class="item_maintenance_header">Dates</td>
                      </tr>
                      <tr>
                        <td colspan="2"><span class="column_divider">0 - Shipment:</span><br />
                          <span class="column_divider">
                          <input type="text" id="txtDateShipment" name="txtDateShipment" maxlength="10" size="10" value="<%= session("date_shipment") %>" />
                          <em>DD/MM/YYYY</em></span></td>
                      </tr>
                      <tr>
                        <td colspan="2">1 - Melb ETA:<br />
                          <input type="text" id="txtMelbETA" name="txtMelbETA" maxlength="10" size="10" value="<%= session("eta_discharged") %>" />
                          <em>DD/MM/YYYY</em></td>
                      </tr>
                      <tr>
                        <td width="50%">2 - Container ETA:<br />
                          <input type="text" id="txtContainerETA" name="txtContainerETA" maxlength="10" size="10" value="<%= session("eta_availability") %>" />
                          <em>DD/MM/YYYY</em></td>
                        <td width="50%">2 - Container (Time):<br />
                          <input type="text" id="txtContainerEtaTime" name="txtContainerEtaTime" maxlength="7" size="10" value="<%= session("melb_eta_time") %>" /></td>
                      </tr>
                      <tr>
                        <td colspan="2">3 - Unpack ETA:<br />
                          <input type="text" id="txtUnpackETA" name="txtUnpackETA" maxlength="10" size="10" value="<%= session("eta_unpacked") %>" />
                          <em>DD/MM/YYYY</em></td>
                      </tr>
                    </table>
                    <br />
                    <table cellpadding="5" cellspacing="0" class="item_maintenance_box">
                      <tr>
                        <td colspan="5" class="item_maintenance_header">Invoices</td>
                      </tr>
                      <tr>
                        <td width="20%"># 1<span class="mandatory">*</span>:<br />
                          <input type="text" id="txtInvoice1" name="txtInvoice1" maxlength="12" size="10" value="<%= session("invoice1") %>" /></td>
                        <td width="20%"># 2:<br />
                          <input type="text" id="txtInvoice2" name="txtInvoice2" maxlength="12" size="10" value="<%= session("invoice2") %>" /></td>
                        <td width="20%"># 3:<br />
                          <input type="text" id="txtInvoice3" name="txtInvoice3" maxlength="12" size="10" value="<%= session("invoice3") %>" /></td>
                        <td width="20%"># 4:<br />
                          <input type="text" id="txtInvoice4" name="txtInvoice4" maxlength="12" size="10" value="<%= session("invoice4") %>" /></td>
                        <td width="20%"># 5:<br />
                          <input type="text" id="txtInvoice5" name="txtInvoice5" maxlength="12" size="10" value="<%= session("invoice5") %>" /></td>
                      </tr>
                    </table>
                    <br />
                    <table cellpadding="5" cellspacing="0" class="item_maintenance_box">
                      <tr>
                        <td colspan="2" class="item_maintenance_header">Documentation</td>
                      </tr>
                      <tr>
                        <td width="50%"><input type="checkbox" name="chkBillLading" id="chkBillLading" value="1" <% if session("bill_lading") = "1" then Response.Write " checked" end if%> />
                          Bill of Lading</td>
                        <td width="50%"><input type="checkbox" name="chkCommercialInvoice" id="chkCommercialInvoice" value="1" <% if session("commercial_invoice") = "1" then Response.Write " checked" end if%> />
                          Commercial Invoice</td>
                      </tr>
                      <tr>
                        <td><input type="checkbox" name="chkPackingList" id="chkPackingList" value="1" <% if session("packing_list") = "1" then Response.Write " checked" end if%> />
                          Packing List</td>
                        <td><input type="checkbox" name="chkPackingDeclaration" id="chkPackingDeclaration" value="1" <% if session("packing_declaration") = "1" then Response.Write " checked" end if%> />
                          Packing Declaration</td>
                      </tr>
                      <tr>
                        <td><input type="checkbox" name="chkTreatmentCertificate" id="chkTreatmentCertificate" value="1" <% if session("treatment_certificate") = "1" then Response.Write " checked" end if%> />
                          Manufacturer Declaration</td>
                        <td><input type="checkbox" name="chkCertificateOrigin" id="chkCertificateOrigin" value="1" <% if session("certificate_origin") = "1" then Response.Write " checked" end if%> />
                          Certificate of Origin</td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td valign="top" colspan="2"><input type="hidden" name="Action" />
                    <input type="submit" value="Update Shipment" <% if session("status") = "0" and session("UsrUserName") <> "jeffj" and session("UsrUserName") <> "yvonnem" and session("UsrUserName") <> "kurtt" and session("UsrUserName") <> "bronb" then Response.Write "disabled" end if%> /></td>
                </tr>
              </table>
            </form></td>
          <td valign="top" style="padding-top:5px; padding-left:5px;"><table cellpadding="5" cellspacing="0" class="fta_box">
              <tr>
                <td colspan="2" class="item_maintenance_header">FTA - <%= session("fta_status") %> <img src="images/icon_new.gif" border="0" align="top" /></td>
              </tr>
              <tr>
                <td colspan="2" class="dotted_line"><form action="" method="post" onsubmit="return submitCertificateOrigin(this)">
                    <input type="checkbox" name="chkCertificateOrigin" id="chkCertificateOrigin" value="1" <% if session("fta_certificate_origin") = "1" then Response.Write " checked" end if%> />
                    <strong><img src="images/bullet_certificate-origin.gif" border="0" /> Certificate of Origin</strong> - <%= displayDateFormatted(session("fta_certificate_origin_date")) %>
                    <div align="left">
                      <input type="hidden" name="Action" />
                      <input type="submit" value="Update Certificate Origin" style="width:230px;" <% if session("fta_certificate_origin") = "1" or session("status") = "0" then Response.Write "disabled" end if%> />
                    </div>
                  </form></td>
              </tr>
              <% if session("fta_status") = "1" or session("fta_certificate_origin") = "1" then %>
              <tr>
                <td colspan="2" class="dotted_line"><form action="" method="post" onsubmit="return submitRefundApplication(this)">
                    <input type="checkbox" name="chkRefundApplication" id="chkRefundApplication" value="1" <% if session("fta_refund_application") = "1" then Response.Write " checked" end if%> />
                    <strong><img src="images/bullet_refund-application.gif" border="0" /> Refund Application</strong> - <%= displayDateFormatted(session("fta_refund_application_date")) %>
                    <div align="left">
                      <input type="hidden" name="Action" />
                      <input type="submit" value="Update Refund Application" style="width:230px;" <% if session("fta_refund_application") = "1" or session("status") = "0" then Response.Write "disabled" end if%> />
                    </div>
                  </form></td>
              </tr>
              <% end if %>
              <% if session("fta_status") = "2" or session("fta_refund_application") = "1" then %>
              <tr>
                <td colspan="2" class="dotted_line"><form action="" method="post" onsubmit="return submitImportDeclaration(this)">
                    <input type="checkbox" name="chkImportDeclaration" id="chkImportDeclaration" value="1" <% if session("fta_import_declaration") = "1" then Response.Write " checked" end if%> />
                    <strong><img src="images/bullet_import-declaration.gif" border="0" /> Import Declaration</strong> - <%= displayDateFormatted(session("fta_import_declaration_date")) %>
                    <div align="left">
                      <input type="hidden" name="Action" />
                      <input type="submit" value="Update Import Declaration" style="width:230px;" <% if session("fta_import_declaration") = "1" or session("status") = "0" then Response.Write "disabled" end if%> />
                    </div>
                  </form></td>
              </tr>
              <% end if %>
              <% if session("fta_status") = "3" or session("fta_import_declaration") = "1" then %>
              <tr>
                <td colspan="2"><form action="" method="post" onsubmit="return submitRefund(this)">
                    <input type="checkbox" name="chkRefund" id="chkRefund" value="1" <% if session("fta_refund") = "1" then Response.Write " checked" end if%> />
                    <strong><img src="images/tick.gif" border="0" /> Refund</strong> - <%= displayDateFormatted(session("fta_refund_date")) %>
                    <div align="left">
                      <input type="hidden" name="Action" />
                      <input type="submit" value="Update Refund" style="width:230px;" <% if session("fta_refund") = "1" or session("status") = "0" then Response.Write "disabled" end if%> />
                    </div>
                  </form></td>
              </tr>
              <% end if %>
            </table>
            <br />
            <table cellpadding="5" cellspacing="0" class="serial_no_box">
              <tr>
                <td class="item_maintenance_header">Documents</td>
              </tr>
              <tr>
                <td><form id="formIDdoc" name="formIDdoc" class="form" method="post">                    
                    <p>
                      <input class="text-input" name="uploadify" id="uploadify" type="file" size="20" />
                    </p>
                    <h3 align="right"><img src="images/forward_arrow.gif" border="0" /> <a href="javascript:Send_document()">Upload files</a></h3>
                    <div id="filesUploaded"></div>
                  </form>
                  <p><% ListFolderContents("\\\YAMMAS22\shipment\") %></p></td>
              </tr>
            </table>
            </td>
        </tr>
      </table>
      <h2>Comments<br />
        <img src="images/comment_bar.jpg" border="0" /></h2>
      <table cellpadding="5" cellspacing="0" border="0" class="comments_box">
        <%= strCommentsList %>
        <tr>
          <td><form action="" method="post" onsubmit="return submitComment(this)">
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