<!--#include file="include/connection_it.asp " -->
<% strSection = "shipment" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Add Shipment</title>
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
	reason += validateEmptyField(theForm.cboDepartment,"Department");
	reason += validateSpecialCharacters(theForm.txtEFT,"EFT");	
	reason += validateSpecialCharacters(theForm.txtVesselName,"Vessel Name");	
	reason += validateSpecialCharacters(theForm.txtVoyage,"Voyage");
	reason += validateEmptyField(theForm.cboTEU,"TEU");	
	reason += validateEmptyField(theForm.txtCartons,"No of Cartons");	
	reason += validateEmptyField(theForm.txtInvoice1,"Invoice 1");
	
  	if (reason != "") {
    	alert("Some fields need correction:\n" + reason);
    	
		blnSubmit = false;
		return false;
  	}

	if (blnSubmit == true){
        theForm.Action.value = 'Add';
		
		return true;
    }
}
</script>
<%
sub addShipment
	dim strSQL
	
	dim strContainer
	dim strSupplierInvoiceNo
	dim strDepartment
	dim strSupplierContact
	dim strCustomCleared
	dim strFumigation
	dim intAirFreight
	dim intPriority
	dim intTransit
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
	dim strMelbETA
	dim strMelbTime
	dim strContainerETA
	dim strContainerTime
	dim strUnpackETA
	dim strUnpackTime
	dim strTEU
	dim strInvoice1
	dim strInvoice2
	dim strInvoice3
	dim strInvoice4
	dim strInvoice5
	dim intPaperwork
	dim strDeliveryType
	
	strContainer 			= trim(Request.Form("txtContainerNo"))
	strSupplierInvoiceNo 	= trim(Request.Form("txtSupplierInvoiceNo"))	
	strDepartment 			= trim(Request.Form("cboDepartment"))
	strSupplierContact 		= trim(Request.Form("cboSupplierContact"))
	strCustomCleared 		= trim(Request.Form("cboCustomCleared"))	
	strFumigation 			= trim(Request.Form("cboFumigation"))	
	intAirFreight 			= trim(Request.Form("chkAirFreight"))
	intPriority 			= trim(Request.Form("chkPriority"))
	intTransit 				= trim(Request.Form("chkTransit"))
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
	strCommodity 			= trim(Request.Form("cboCommodity"))	
	strPortOrigin 			= trim(Request.Form("cboPortOrigin"))	
	strCountryOrigin 		= trim(Request.Form("cboCountryOrigin"))	
	strVesselName 			= trim(Request.Form("txtVesselName"))	
	strVoyage 				= trim(Request.Form("txtVoyage"))	
	strWarehouse 			= trim(Request.Form("cboWarehouse"))
	intCartons 				= trim(Request.Form("txtCartons"))
	strDateShipment 		= trim(Request.Form("txtDateShipment"))	
	strMelbETA 				= trim(Request.Form("txtMelbETA"))
	strMelbTime 			= trim(Request.Form("txtMelbTime"))
	strContainerETA 		= trim(Request.Form("txtContainerETA"))	
	strContainerTime 		= trim(Request.Form("txtContainerEtaTime"))
	strUnpackETA 			= trim(Request.Form("txtUnpackETA"))
	strUnpackTime 			= trim(Request.Form("txtUnpackTime"))
	strTEU 					= trim(Request.Form("cboTEU"))	
	strInvoice1 			= trim(Request.Form("txtInvoice1"))
	strInvoice2 			= trim(Request.Form("txtInvoice2"))
	strInvoice3 			= trim(Request.Form("txtInvoice3"))
	strInvoice4 			= trim(Request.Form("txtInvoice4"))
	strInvoice5 			= trim(Request.Form("txtInvoice5"))
	intPaperwork			= trim(Request.Form("chkPaperwork")) 
	strDeliveryType 		= trim(Request.Form("cboDeliveryType"))
	
	call OpenDataBase()
		
	strSQL = "INSERT INTO yma_shipment ("
	strSQL = strSQL & "container_no, "
	strSQL = strSQL & "supplier_invoice_no, "
	strSQL = strSQL & "department, "
	strSQL = strSQL & "supplier_contact, "
	strSQL = strSQL & "custom_cleared, "
	strSQL = strSQL & "fumigation, "
	strSQL = strSQL & "air_freight, "
	strSQL = strSQL & "priority, "
	strSQL = strSQL & "transit, "
	strSQL = strSQL & "edo, "
	strSQL = strSQL & "FF, "
	strSQL = strSQL & "EFT, "
	strSQL = strSQL & "all_documents, "
	strSQL = strSQL & "bill_lading, "
	strSQL = strSQL & "commercial_invoice, "
	strSQL = strSQL & "packing_list, "
	strSQL = strSQL & "packing_declaration, "
	strSQL = strSQL & "treatment_certificate, "
	strSQL = strSQL & "certificate_origin, "
	strSQL = strSQL & "comments, "
	strSQL = strSQL & "commodity, "
	strSQL = strSQL & "port_origin, "
	strSQL = strSQL & "country_origin, "
	strSQL = strSQL & "vessel_name, "
	strSQL = strSQL & "voyage, "
	strSQL = strSQL & "warehouse, "
	strSQL = strSQL & "cartons, "
	strSQL = strSQL & "date_shipment, "
	strSQL = strSQL & "eta_discharged, "
	strSQL = strSQL & "melb_time, "
	strSQL = strSQL & "eta_availability, "
	strSQL = strSQL & "melb_eta_time, "
	strSQL = strSQL & "eta_unpacked, "
	strSQL = strSQL & "unpack_time, "
	strSQL = strSQL & "teu, "
	strSQL = strSQL & "paperwork, "
	strSQL = strSQL & "delivery_type, "
	strSQL = strSQL & "created_by, "
	strSQL = strSQL & "invoice1, invoice2, invoice3, invoice4, invoice5"
	strSQL = strSQL & ") VALUES ( "
	strSQL = strSQL & "'" & strContainer & "',"	
	strSQL = strSQL & "'" & strSupplierInvoiceNo & "',"	
	strSQL = strSQL & "'" & strDepartment & "',"
	strSQL = strSQL & "'" & strSupplierContact & "',"
	strSQL = strSQL & "'" & strCustomCleared & "',"	
	strSQL = strSQL & "'" & strFumigation & "',"	
	strSQL = strSQL & "'" & intAirFreight & "',"
	strSQL = strSQL & "'" & intPriority & "',"
	strSQL = strSQL & "'" & intTransit & "',"
	strSQL = strSQL & "'" & intEDO & "',"
	strSQL = strSQL & "'" & strFF & "',"
	strSQL = strSQL & "'" & strEFT & "',"
	strSQL = strSQL & "'" & strAllDocuments & "',"
	strSQL = strSQL & "'" & intBillLading & "',"
	strSQL = strSQL & "'" & intCommercialInvoice & "',"
	strSQL = strSQL & "'" & intPackingList & "',"
	strSQL = strSQL & "'" & intPackingDeclaration & "',"
	strSQL = strSQL & "'" & intTreatmentCertificate & "',"
	strSQL = strSQL & "'" & intCertificateOrigin & "',"
	strSQL = strSQL & "'" & Server.HTMLEncode(strComments) & "',"
	strSQL = strSQL & "'" & strCommodity & "',"	
	strSQL = strSQL & "'" & strPortOrigin & "',"	
	strSQL = strSQL & "'" & strCountryOrigin & "',"	
	strSQL = strSQL & "'" & strVesselName & "',"	
	strSQL = strSQL & "'" & strVoyage & "',"
	strSQL = strSQL & "'" & strWarehouse & "',"
	strSQL = strSQL & "'" & intCartons & "',"	
	strSQL = strSQL & " CONVERT(datetime,'" & strDateShipment & "',103),"
	strSQL = strSQL & " CONVERT(datetime,'" & strMelbETA & "',103),"
	strSQL = strSQL & "'" & strMelbTime & "',"
	strSQL = strSQL & " CONVERT(datetime,'" & strContainerETA & "',103),"	
	strSQL = strSQL & "'" & strContainerTime & "',"
	strSQL = strSQL & " CONVERT(datetime,'" & strUnpackETA & "',103),"
	strSQL = strSQL & "'" & strUnpackTime & "',"
	strSQL = strSQL & "'" & strTEU & "',"
	strSQL = strSQL & "'" & intPaperwork & "',"
	strSQL = strSQL & "'" & strDeliveryType & "',"
	strSQL = strSQL & "'" & session("UsrUserName") & "',"
	strSQL = strSQL & "'" & strInvoice1 & "', '" & strInvoice2 & "', '" & strInvoice3 & "', '" & strInvoice4 & "', '" & strInvoice5 & "')"
	
	'response.Write strSQL 
	on error resume next
	conn.Execute strSQL
		
	if err <> 0 then
		strMessageText = err.description
	else 	
		Response.Redirect("list_shipment.asp")
	end if 

	call CloseDataBase()
end sub

sub main
	call UTL_validateLogin 
	
	if Request.ServerVariables("REQUEST_METHOD") = "POST" then
		if Trim(Request("Action")) = "Add" then
			call addShipment
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
          <td><a href="list_shipment.asp"><img src="images/icon_shipment.jpg" border="0" alt="Shipment" /></a></td>
          <td valign="top"><img src="images/backward_arrow.gif" border="0" /> <a href="list_shipment.asp">Back to List</a>
            <h2>Add Shipment</h2>
            <font color="green"><%= strMessageText %></font></td>
        </tr>
      </table>
      <form action="" method="post" name="form_add_shipment" id="form_add_shipment" onsubmit="return validateFormOnSubmit(this)">
        <table border="0" cellpadding="5" cellspacing="0" class="wide_table">
          <tr>
            <td width="33%" valign="top"><table cellpadding="5" cellspacing="0" class="item_maintenance_box">
                <tr>
                  <td colspan="3" class="item_maintenance_header">Container Info</td>
                </tr>
                <tr>
                  <td>Container no<span class="mandatory">*</span>:<br />
                    <input type="text" id="txtContainerNo" name="txtContainerNo" maxlength="20" size="20" />
                    <em>(TBA if blank)</em></td>
                  <td colspan="2">Supplier invoice no<span class="mandatory">*</span>:<br />
                    <input type="text" id="txtSupplierInvoiceNo" name="txtSupplierInvoiceNo" maxlength="20" size="20" /></td>
                </tr>
                <tr>
                  <td>Department<span class="mandatory">*</span>:<br />
                    <select name="cboDepartment">
                      <option value="">...</option>
                      <option value="AV">AV</option>
                      <option value="MPD">MPD</option>
                      <option value="AV-MPD">AV &amp; MPD</option>
                      <option value="Other">Other</option>
                    </select></td>
                  <td colspan="2">Supplier contact:<br />
                    <select name="cboSupplierContact">
                      <option value="">...</option>
                      <option value="distorder@steinberg.de">Steinberg (Annegret)</option>
                      <option value="jwigger@paiste.com">Paiste (Jasmin)</option>
                      <option value="mark@belcat.com">Belcat (Mark)</option>
                      <option value="takaoka@korg.co.jp">Korg (Shigeru)</option>
                      <option value="fujiwara@korg.co.jp">Korg (Shungo)</option>
                      <option value="mamii@korg.jp">Korg (Naoko)</option>
                      <option value="khsms1@khs-musix.com">KHS (Iris)</option>
                      <option value="msorensen@yamaha.com">YAMUSA (Mellonie)</option>
                      <option value="ismerdel@yamaha.com">Everbright (Irene)</option>
                    </select></td>
                </tr>
                <tr>
                  <td><input type="checkbox" name="chkAirFreight" id="chkAirFreight" value="1" />
                    <label for="chkAirFreight">Air freight</label> <img src="images/airplane.gif" alt="" border="0" /></td>
                  <td colspan="2"><input type="checkbox" name="chkPriority" id="chkPriority" value="1" />
                    <label for="chkPriority">Priority</label> <img src="images/icon_priority.gif" alt="" border="0" /></td>
                </tr>
                <tr>
                  <td><input type="checkbox" name="chkTransit" id="chkTransit" value="1" />
                    <label for="chkTransit">In Transit</label> <img src="images/icon_new.gif" border="0" align="top" /></td>
                  <td colspan="2">&nbsp;</td>
                </tr>
                <tr>
                  <td width="50%">Custom cleared:<br />
                    <select name="cboCustomCleared">
                      <option value="-">...</option>
                      <option value="Y">Yes</option>
                      <option value="N">No</option>
                      <option value="Border-hold">Border hold</option>
                      <option value="X-ray">X-ray</option>
                    </select></td>
                  <td width="20%">Fumigation:<br />
                    <select name="cboFumigation">
                      <option value="-">...</option>
                      <option value="Y">Yes</option>
                      <option value="N">No</option>
                      <option value="AQIS inspect">AQIS inspect</option>
                    </select></td>
                  <td width="30%">Freight fwd:<br />
                    <select name="cboFF">
                      <option value="-">...</option>
                      <option value="M">M</option>
                      <option value="K">K</option>
                      <option value="Y">Y</option>
                    </select></td>
                </tr>
                <tr>
                  <td>EFT:<br />
                    <input type="text" id="txtEFT" name="txtEFT" maxlength="6" size="6" /></td>
                  <td colspan="2">All Documents:<br />
                    <select name="cboAllDocuments">
                      <option value="N">No</option>
                      <option value="Y">Yes</option>
                      <option value="Part">Part</option>
                    </select></td>
                </tr>
                <tr>
                  <td>Commodity:<br />
                    <select name="cboCommodity">
                      <option value="NA">...</option>
                      <option value="Amps">Amps</option>
                      <option value="AV Products">AV Products</option>
                      <option value="AV Receivers">AV Receivers</option>
                      <option value="AV Speakers">AV Speakers</option>
                      <option value="Blu-ray Players">Blu-ray Players</option>
                      <option value="Cymbals">Cymbals</option>
                      <option value="Drums Kit">Drums Kit</option>
                      <option value="Electric Drums">Electric Drums</option>
                      <option value="Electric Pianos and Mixers">Electric Pianos and Mixers</option>
                      <option value="Guitars">Guitars</option>
                      <option value="Guitars and Drums">Guitars and Drums</option>
                      <option value="Keyboards">Keyboards</option>
                      <option value="Keyboard Stands">Keyboard Stands</option>
                      <option value="Misc MPD">Misc MPD</option>
                      <option value="MPD Speakers">MPD Speakers</option>
                      <option value="Pianos">Pianos</option>
                      <option value="Returned goods">Returned goods</option>
                      <option value="Speakers">Speakers</option>
                      <option value="VOX">VOX</option>
                      <option value="AV and MPD Misc">AV and MPD Misc</option>
                      <option value="Line 6">Line 6</option>
                      <option value="Other">Other</option>
                    </select></td>
                  <td colspan="2">TEU<span class="mandatory">*</span>:<br />
                    <select name="cboTEU">
                      <option value="">...</option>
                      <option value="20">20</option>
                      <option value="40">40</option>
                      <option value="HC">HC</option>
                      <option value="LCL">LCL</option>
                      <option value="Air Freight">Air Freight</option>
                    </select></td>
                </tr>
                <tr>
                  <td>Port of origin:<br />
                    <select name="cboPortOrigin">
                      <option value="NA">...</option>
                      <option value="Auckland">Auckland</option>
                      <option value="Dalian">Dalian</option>
                      <option value="Hamburg">Hamburg</option>
                      <option value="Ho Chi Minh">Ho Chi Minh</option>
                      <option value="Hongkong">Hongkong</option>
                      <option value="Jakarta">Jakarta</option>
                      <option value="Lithia Springs">Lithia Springs</option>
                      <option value="Port Klang">Port Klang</option>
                      <option value="Savannah">Savannah</option>
                      <option value="Semarang">Semarang</option>
                      <option value="Shanghai">Shanghai</option>
                      <option value="Shekou">Shekou</option>
                      <option value="Shimizu">Shimizu</option>
                      <option value="Singapore">Singapore</option>
                      <option value="Surabaya">Surabaya</option>
                      <option value="Thames">Thames</option>
                      <option value="Xingang">Xingang</option>
                      <option value="Yantian">Yantian</option>
                      <option value="Yokohama">Yokohama</option>
                      <option value="Other">Other</option>
                    </select></td>
                  <td colspan="2">Country of origin:<br />
                    <select name="cboCountryOrigin">
                      <option value="NA">...</option>
                      <option value="China">China</option>
                      <option value="England">England</option>
                      <option value="Germany">Germany</option>
                      <option value="Hongkong">Hongkong</option>
                      <option value="Indonesia">Indonesia</option>
                      <option value="Japan">Japan</option>
                      <option value="Malaysia">Malaysia</option>
                      <option value="NZ">NZ</option>
                      <option value="Singapore">Singapore</option>
                      <option value="USA">USA</option>
                      <option value="Vietnam">Vietnam</option>
                      <option value="Other">Other</option>
                    </select></td>
                </tr>
                <tr>
                  <td>Vessel:<br />
                    <input type="text" id="txtVesselName" name="txtVesselName" maxlength="30" size="30" /></td>
                  <td colspan="2">Voyage:<br />
                    <input type="text" id="txtVoyage" name="txtVoyage" maxlength="8" size="8" /></td>
                </tr>
                <tr>
                  <td>Warehouse: <br />
                    <select name="cboWarehouse">
                      <option value="TT">TT</option>
                      <option value="EXL">EXL</option>
                      <option value="YMA">YMA Head Office</option>
                    </select></td>
                  <td colspan="2">No of cartons<span class="mandatory">*</span>:<br />
                    <input type="text" id="txtCartons" name="txtCartons" maxlength="5" size="5" /></td>
                </tr>
                <tr>
                  <td colspan="3"><input type="checkbox" name="chkEDO" id="chkEDO" value="1" /> <label for="chkEDO">EDO</label></td>                  
                </tr>
                <tr>
                  <td valign="top"><input type="checkbox" name="chkPaperwork" id="chkPaperwork" value="1" /> <label for="chkPaperwork">Paperwork sent to Rocke</label> <img src="images/icon_new.gif" border="0" align="top" /></td>
                  <td colspan="2">Delivery Type: <img src="images/icon_new.gif" border="0" align="top" /><br />
                         <select name="cboDeliveryType">
                            <option value="TBA">TBA</option>
                            <option value="Normal">Normal</option>
                            <option value="Drop out">Drop out</option>                            
                          </select></td>
                </tr>
                <tr>
                  <td colspan="3"></td>
                </tr>
              </table></td>
            <td width="33%" valign="top"><table cellpadding="5" cellspacing="0" class="item_maintenance_box">
                <tr>
                  <td colspan="2" class="item_maintenance_header">Dates</td>
                </tr>
                <tr>
                  <td colspan="2"><span class="column_divider">0 - Shipment:</span><br />
                    <span class="column_divider">
                    <input type="text" id="txtDateShipment" name="txtDateShipment" maxlength="10" size="20" /></span></td>
                </tr>
                <tr>
                  <td>1 - Melb ETA:<br />
                    <input type="text" id="txtMelbETA" name="txtMelbETA" maxlength="10" size="10" />
                    <em>DD/MM/YYYY</em></td>
                  <td>1. Melb ETA Time:<br />
                  <input type="text" id="txtMelbTime" name="txtMelbTime" maxlength="7" size="10" /></td>
                </tr>
                <tr>
                  <td width="50%">2 - Container ETA:<br />
                    <input type="text" id="txtContainerETA" name="txtContainerETA" maxlength="10" size="10" />
                    <em>DD/MM/YYYY</em></td>
                  <td width="50%">2 - Container ETA Time:<br />
                    <input type="text" id="txtContainerEtaTime" name="txtContainerEtaTime" maxlength="7" size="10" /></td>
                </tr>
                <tr>
                  <td>3 - Unpack ETA:<br />
                    <input type="text" id="txtUnpackETA" name="txtUnpackETA" maxlength="10" size="10" />
                    <em>DD/MM/YYYY</em></td>
                  <td>3. Unpack ETA Time:<br />
                  <input type="text" id="txtUnpackTime" name="txtUnpackTime" maxlength="7" size="10" /></td>
                </tr>
              </table>
              <br />
              <table cellpadding="5" cellspacing="0" class="item_maintenance_box">
                <tr>
                  <td colspan="5" class="item_maintenance_header">Invoices</td>
                </tr>
                <tr>
                  <td width="20%"># 1<span class="mandatory">*</span>:<br />
                    <input type="text" id="txtInvoice1" name="txtInvoice1" maxlength="12" size="8" /></td>
                  <td width="20%"># 2:<br />
                    <input type="text" id="txtInvoice2" name="txtInvoice2" maxlength="12" size="8" /></td>
                  <td width="20%"># 3:<br />
                    <input type="text" id="txtInvoice3" name="txtInvoice3" maxlength="12" size="8" /></td>
                  <td width="20%"># 4:<br />
                    <input type="text" id="txtInvoice4" name="txtInvoice4" maxlength="12" size="8" /></td>
                  <td width="20%"># 5:<br />
                    <input type="text" id="txtInvoice5" name="txtInvoice5" maxlength="12" size="8" /></td>
                </tr>
              </table>
              <br />
              <table cellpadding="5" cellspacing="0" class="item_maintenance_box">
                <tr>
                  <td colspan="2" class="item_maintenance_header">Documentation</td>
                </tr>
                <tr>
                  <td width="50%"><input type="checkbox" name="chkBillLading" id="chkBillLading" value="1" />
                    <label for="chkBillLading">Bill of Lading</label></td>
                  <td width="50%"><input type="checkbox" name="chkCommercialInvoice" id="chkCommercialInvoice" value="1" />
                    <label for="chkCommercialInvoice">Commercial Invoice</label></td>
                </tr>
                <tr>
                  <td><input type="checkbox" name="chkPackingList" id="chkPackingList" value="1" />
                    <label for="chkPackingList">Packing List</label></td>
                  <td><input type="checkbox" name="chkPackingDeclaration" id="chkPackingDeclaration" value="1" />
                    <label for="chkPackingDeclaration">Packing Declaration</label></td>
                </tr>
                <tr>
                  <td><input type="checkbox" name="chkTreatmentCertificate" id="chkTreatmentCertificate" value="1" />
                    <label for="chkTreamentCertificate">Manufacturer Declaration</label></td>
                  <td><input type="checkbox" name="chkCertificateOrigin" id="chkCertificateOrigin" value="1" />
                   <label for="chkCertificateOrigin">Certificate of Origin</label></td>
                </tr>
              </table></td>
            <td width="33%" valign="top"><table cellpadding="5" cellspacing="0" class="item_maintenance_box">
                <tr>
                  <td colspan="2" class="item_maintenance_header">Additional Info</td>
                </tr>
                <tr>
                  <td colspan="2"><textarea name="txtComments" id="txtComments" cols="55" rows="3"></textarea></td>
                </tr>
              </table></td>
          </tr>
        </table>
          <input type="hidden" name="Action" />
          <input type="submit" value="Add Shipment" />
      </form></td>
  </tr>
</table>
<script type="text/javascript" src="include/moment.js"></script>
<script type="text/javascript" src="include/pikaday.js"></script>
<script type="text/javascript">	
	var picker = new Pikaday(
    {
        field: document.getElementById('txtDateShipment'),		
        firstDay: 1,
        minDate: new Date('2013-01-01'),
        maxDate: new Date('2020-12-31'),
        yearRange: [2013,2020],
		format: 'DD/MM/YYYY'
    });
	
	var picker = new Pikaday(
    {
        field: document.getElementById('txtMelbETA'),		
        firstDay: 1,
        minDate: new Date('2014-01-01'),
        maxDate: new Date('2020-12-31'),
        yearRange: [2014,2020],
		format: 'DD/MM/YYYY'
    });
	
	var picker = new Pikaday(
    {
        field: document.getElementById('txtContainerETA'),		
        firstDay: 1,
        minDate: new Date('2014-01-01'),
        maxDate: new Date('2020-12-31'),
        yearRange: [2014,2020],
		format: 'DD/MM/YYYY'
    });
	
	var picker = new Pikaday(
    {
        field: document.getElementById('txtUnpackETA'),		
        firstDay: 1,
        minDate: new Date('2014-01-01'),
        maxDate: new Date('2020-12-31'),
        yearRange: [2014,2020],
		format: 'DD/MM/YYYY'
    });
	
</script>
</body>
</html>