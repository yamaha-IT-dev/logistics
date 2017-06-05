<!--#include file="include/connection_it.asp " -->
<!--#include file="class/clsItemMaintenance.asp " -->
<%
Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<% strSection = "item" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<script type="text/javascript" src="include/usableforms.js"></script>
<script type="text/javascript" src="include/jquery.min.js"></script>
<script type="text/javascript" src="include/ddaccordion.js">
/***********************************************
* Accordion Content script- (c) Dynamic Drive DHTML code library (www.dynamicdrive.com)
* Visit http://www.dynamicDrive.com for hundreds of DHTML scripts
* This notice must stay intact for legal use
***********************************************/
</script>
<style type="text/css">
.myheadings { /*header of 1st demo*/
    cursor: hand;
    cursor: pointer;
    font-size:12px;
    /*border: 1px solid gray;*/
    background: #E1E1E1;
}

.openpet { /*class added to contents of 1st demo when they are open*/
    background:#CCC;
    color:#999;
}

#loading-spinner {
    background: url("images/loading.gif") center center no-repeat;
    height: 100%;
    z-index: 20;
}

.overlay {
    background: #FFFFFF;
    display: none;
    position: absolute;
    top: 0;
    right: 0;
    bottom: 0;
    left: 0;
    opacity: 0.75;
}
</style>
<script type="text/javascript">
//Initialize:
ddaccordion.init({
    headerclass: "myheadings", //Shared CSS class name of headers group
    contentclass: "thecontent", //Shared CSS class name of contents group
    revealtype: "click", //Reveal content when user clicks or onmouseover the header? Valid value: "click", "clickgo", or "mouseover"
    mouseoverdelay: 200, //if revealtype="mouseover", set delay in milliseconds before header expands onMouseover
    collapseprev: false, //Collapse previous content (so only one open at any time)? true/false 
    defaultexpanded: [0], //index of content(s) open by default [index1, index2, etc]. [] denotes no content.
    onemustopen: false, //Specify whether at least one header should be open always (so never all headers closed)
    animatedefault: false, //Should contents open by default be animated into view?
    persiststate: true, //persist state of opened contents within browser session?
    toggleclass: ["", "openpet"], //Two CSS classes to be applied to the header when it's collapsed and expanded, respectively ["class1", "class2"]
    togglehtml: ["none", "", ""], //Additional HTML added to the header when it's collapsed and expanded, respectively  ["position", "html1", "html2"] (see docs)
    animatespeed: "fast", //speed of animation: integer in milliseconds (ie: 200), or keywords "fast", "normal", or "slow"
    oninit:function(expandedindices){ //custom code to run when headers have initalized
        //do nothing
    },
    onopenclose:function(header, index, state, isuseractivated){ //custom code to run whenever a header is opened or closed
        //do nothing
    }
})
</script>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Add Item Maintenance</title>
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="include/usableforms.js"></script>
<script type="text/javascript" src="include/generic_form_validations.js"></script>
<script type="text/javascript" src="include/javascript.js"></script>
<script language="JavaScript" type="text/javascript">
function validateItemMaintenanceForm(theForm) {
    var reason = "";
    var blnSubmit = true;

    reason += validateEmptyField(theForm.cboDepartment,"Department");
    reason += validateEmptyField(theForm.txtBaseCode,"BASE Code");
    reason += validateSpecialCharacters(theForm.txtBaseCode,"BASE Code");
    reason += validateEmptyField(theForm.txtItemName,"Item Name");
    reason += validateSpecialCharacters(theForm.txtItemName,"Item Name");
    reason += validateEmptyField(theForm.txtModelName,"Model Name");
    reason += validateSpecialCharacters(theForm.txtModelName,"Model Name");
    reason += validateEmptyField(theForm.txtDescription,"Description");
    reason += validateSpecialCharacters(theForm.txtDescription,"Description");
    reason += validateEmptyField(theForm.txtGMCcode,"GMC Code");
    reason += validateSpecialCharacters(theForm.txtGMCcode,"GMC Code");

    if (theForm.cboMulticolour.value != 0) {
        reason += validateEmptyField(theForm.txtColour1BaseCode,"Multicolour 1: BASE Code");
        reason += validateEmptyField(theForm.txtColour1,"Multicolour 1: Colour");
        reason += validateEmptyField(theForm.txtGMC1,"Multicolour 1: GMC Code");
        reason += validateEmptyField(theForm.txtEAN1,"Multicolour 1: EAN Code");
    }

    if (theForm.cboSetItem.value != 0) {
        reason += validateEmptyField(theForm.txtItem1,"Set Item 1: BASE Code");
        reason += validateEmptyField(theForm.txtQty1,"Set Item 1: Qty");
    }

    reason += validateEmptyField(theForm.txtGrossWeight,"Gross Weight");
    reason += validateSpecialCharacters(theForm.txtGrossWeight,"Gross Weight");
    reason += validateEmptyField(theForm.txtNettWeight,"Nett Weight");
    reason += validateSpecialCharacters(theForm.txtNettWeight,"Nett Weight");
    reason += validateNumeric(theForm.txtWidth,"Width");
    reason += validateNumeric(theForm.txtHeight,"Height");
    reason += validateNumeric(theForm.txtDepth,"Depth");
    reason += validateEmptyField(theForm.txtEANcode,"EAN Code");
    reason += validateEmptyField(theForm.txtTrade,"Trade Pricing");
    reason += validateSpecialCharacters(theForm.txtTrade,"Trade Pricing");
    reason += validateEmptyField(theForm.txtRRP,"RRP");
    reason += validateSpecialCharacters(theForm.txtRRP,"RRP");	

    if (reason != "") {
        alert("Oops... some fields need correction:\n" + reason);

        blnSubmit = false;
        return false;
    }

    if (blnSubmit == true){
        theForm.Action.value = 'Add';

        return true;
    }
}

function YCJItemSet() {
    var strStatus = document.forms[0].cboYCJItem.value;

    if (strStatus != '') {
        // Disply a loading spinner
        $(".overlay").show();

        document.location.href = 'add_item-maintenance.asp?ycjitem='+ strStatus +'&cboYCJItem=' + strStatus;
    }
}
</script>
<%
sub addItemMaintenance
    dim strSQL

    dim strDepartment
    dim strBaseCode
    dim strItemName
    dim strModelName
    dim strDescription
    dim strGMC
    dim strSalesGroup
    dim strShippingGroup
    dim strAccountGroup
    dim strDiscountGroup
    dim intMulticolour
    dim strColour1
    dim strColour1Code
    dim strColour1GMC
    dim strColour1EAN
    dim strColour2
    dim strColour2Code
    dim strColour2GMC
    dim strColour2EAN
    dim strColour3
    dim strColour3Code
    dim strColour3GMC
    dim strColour3EAN
    dim strColour4
    dim strColour4Code
    dim strColour4GMC
    dim strColour4EAN
    dim strColour5
    dim strColour5Code
    dim strColour5GMC
    dim strColour5EAN
    dim strLifecycle
    dim strLifecycleDate
    dim strLifecycleDiscontinuedItem
    dim strLifecycleDiscontinuedDate
    dim strMinOrderQty
    dim strOrderLot
    dim intSerialised
    dim intSetItem
    dim strSet1
    dim strSet1Colour
    dim strSet1Qty
    dim strSet2
    dim strSet2Colour
    dim strSet2Qty
    dim strSet3
    dim strSet3Colour
    dim strSet3Qty
    dim strSet4
    dim strSet4Colour
    dim strSet4Qty
    dim strSet5
    dim strSet5Colour
    dim strSet5Qty
    dim strSet6
    dim strSet6Colour
    dim strSet6Qty
    dim strSet7
    dim strSet7Colour
    dim strSet7Qty
    dim strSet8
    dim strSet8Colour
    dim strSet8Qty
    dim strSet9
    dim strSet9Colour
    dim strSet9Qty
    dim strSet10
    dim strSet10Colour
    dim strSet10Qty
    dim strGrossWeight
    dim strNettWeight
    dim strWidth
    dim strHeight
    dim strDepth
    dim strVolume
    dim strPackUnit
    dim strTariffCode
    dim strVendor
    dim strVendorOther
    dim strCountryOrigin
    dim strEanCode
    dim strFOB
    dim strTrade
    dim strRRP
    dim strNzTrade
	dim S03
	dim S04
	dim S08
    dim intModRequired
    dim strComments

    strDepartment                   = Request.Form("cboDepartment")
    strBaseCode                     = Replace(Request.Form("txtBaseCode"),"'","''")
    strItemName                     = Replace(Request.Form("txtItemName"),"'","''")
    strModelName                    = Replace(Request.Form("txtModelName"),"'","''")
    strDescription                  = Replace(Request.Form("txtDescription"),"'","''")
    strGMC                          = Replace(Request.Form("txtGMCcode"),"'","''")
    strSalesGroup                   = trim(Request.Form("cboSalesGroup"))
    strShippingGroup                = trim(Request.Form("cboShippingGroup"))
    strAccountGroup                 = trim(Request.Form("cboAccountGroup"))
    strDiscountGroup                = trim(Request.Form("cboDiscountGroup"))
    intMulticolour                  = Request.Form("cboMultiColour")
    strColour1                      = Replace(Request.Form("txtColour1"),"'","''")
    strColour1Code                  = Replace(Request.Form("txtColour1BaseCode"),"'","''")
    strColour1GMC                   = Replace(Request.Form("txtGMC1"),"'","''")
    strColour1EAN                   = Replace(Request.Form("txtEAN1"),"'","''")
    strColour2                      = Replace(Request.Form("txtColour2"),"'","''")
    strColour2Code                  = Replace(Request.Form("txtColour2BaseCode"),"'","''")
    strColour2GMC                   = Replace(Request.Form("txtGMC2"),"'","''")
    strColour2EAN                   = Replace(Request.Form("txtEAN2"),"'","''")
    strColour3                      = Replace(Request.Form("txtColour3"),"'","''")
    strColour3Code                  = Replace(Request.Form("txtColour3BaseCode"),"'","''")
    strColour3GMC                   = Replace(Request.Form("txtGMC3"),"'","''")
    strColour3EAN                   = Replace(Request.Form("txtEAN3"),"'","''")
    strColour4                      = Replace(Request.Form("txtColour4"),"'","''")
    strColour4Code                  = Replace(Request.Form("txtColour4BaseCode"),"'","''")
    strColour4GMC                   = Replace(Request.Form("txtGMC4"),"'","''")
    strColour4EAN                   = Replace(Request.Form("txtEAN4"),"'","''")
    strColour5                      = Replace(Request.Form("txtColour5"),"'","''")
    strColour5Code                  = Replace(Request.Form("txtColour5BaseCode"),"'","''")
    strColour5GMC                   = Replace(Request.Form("txtGMC5"),"'","''")
    strColour5EAN                   = Replace(Request.Form("txtEAN5"),"'","''")
    strLifecycle                    = Request.Form("cboLifecycle")
    strLifecycleDate                = Replace(Request.Form("txtLifecycleDate"),"'","''")
    strLifecycleDiscontinuedItem    = Replace(Request.Form("txtAlternativeItem"),"'","''")
    strLifecycleDiscontinuedDate    = Replace(Request.Form("txtAlternativeItemDate"),"'","''")
    strMinOrderQty                  = Replace(Request.Form("txtMinOrderQty"),"'","''")
    strOrderLot                     = Replace(Request.Form("txtOrderLot"),"'","''")
    intSerialised                   = Request.Form("chkSerialised")
    intSetItem                      = Request.Form("cboSetItem")
    strSet1                         = Replace(Request.Form("txtItem1"),"'","''")
    strSet1Colour                   = Request.Form("cboItem1Colour")
    strSet1Qty                      = Replace(Request.Form("txtQty1"),"'","''")
    strSet2                         = Replace(Request.Form("txtItem2"),"'","''")
    strSet2Colour                   = Request.Form("cboItem2Colour")
    strSet2Qty                      = Replace(Request.Form("txtQty2"),"'","''")
    strSet3                         = Replace(Request.Form("txtItem3"),"'","''")
    strSet3Colour                   = Request.Form("cboItem3Colour")
    strSet3Qty                      = Replace(Request.Form("txtQty3"),"'","''")
    strSet4                         = Replace(Request.Form("txtItem4"),"'","''")
    strSet4Colour                   = Request.Form("cboItem4Colour")
    strSet4Qty                      = Replace(Request.Form("txtQty4"),"'","''")
    strSet5                         = Replace(Request.Form("txtItem5"),"'","''")
    strSet5Colour                   = Request.Form("cboItem5Colour")
    strSet5Qty                      = Replace(Request.Form("txtQty5"),"'","''")
    strSet6                         = Replace(Request.Form("txtItem6"),"'","''")
    strSet6Colour                   = Request.Form("cboItem6Colour")
    strSet6Qty                      = Replace(Request.Form("txtQty6"),"'","''")
    strSet7                         = Replace(Request.Form("txtItem7"),"'","''")
    strSet7Colour                   = Request.Form("cboItem7Colour")
    strSet7Qty                      = Replace(Request.Form("txtQty7"),"'","''")
    strSet8                         = Replace(Request.Form("txtItem8"),"'","''")
    strSet8Colour                   = Request.Form("cboItem8Colour")
    strSet8Qty                      = Replace(Request.Form("txtQty8"),"'","''")
    strSet9                         = Replace(Request.Form("txtItem9"),"'","''")
    strSet9Colour                   = Request.Form("cboItem9Colour")
    strSet9Qty                      = Replace(Request.Form("txtQty9"),"'","''")
    strSet10                        = Replace(Request.Form("txtItem10"),"'","''")
    strSet10Colour                  = Request.Form("cboItem10Colour")
    strSet10Qty                     = Replace(Request.Form("txtQty10"),"'","''")
    strGrossWeight                  = Replace(Request.Form("txtGrossWeight"),"'","''")
    strNettWeight                   = Replace(Request.Form("txtNettWeight"),"'","''")
    strWidth                        = Replace(Request.Form("txtWidth"),"'","''")
    strHeight                       = Replace(Request.Form("txtHeight"),"'","''")
    strDepth                        = Replace(Request.Form("txtDepth"),"'","''")
    strVolume                       = (strWidth * strHeight * strDepth) / 1000000
    strPackUnit                     = Replace(Request.Form("txtPackUnit"),"'","''")
    strTariffCode                   = Request.Form("cboTariffCode")
    strVendor                       = Replace(Request.Form("txtVendor"),"'","''")
    strCountryOrigin                = Request.Form("cboCountryOrigin")
    strEanCode                      = Replace(Request.Form("txtEANcode"),"'","''")
    strFOB                          = Replace(Request.Form("txtFOB"),"'","''")
    strTrade                        = Replace(Request.Form("txtTrade"),"'","''")
    strRRP                          = Replace(Request.Form("txtRRP"),"'","''")
    strNzTrade                      = Replace(Request.Form("txtNZtrade"),"'","''")
	strS03                     		= Replace(Request.Form("txtS03"),"'","''")
	strS04                      	= Replace(Request.Form("txtS04"),"'","''")
	strS08                      	= Replace(Request.Form("txtS08"),"'","''")
    intModRequired                  = Request.Form("cboModRequired")
    strComments                     = Replace(Request.Form("txtComments"),"'","''")

    if strColour1Code = "" then
        strColour1GMC = ""
        strColour1EAN = ""
    end if

    if strColour2Code = "" then
        strColour2GMC = ""
        strColour2EAN = ""
    end if

    if strColour3Code = "" then
        strColour3GMC = ""
        strColour3EAN = ""
    end if

    if strColour4Code = "" then
        strColour4GMC = ""
        strColour4EAN = ""
    end if

    if strColour5Code = "" then
        strColour5GMC = ""
        strColour5EAN = ""
    end if

    call OpenDataBase()

    strSQL = "INSERT INTO yma_item_maintenance ("
    strSQL = strSQL & " department, "
    strSQL = strSQL & " base_code, item_name, model_name, description, "
    strSQL = strSQL & " gmc_code, "
    strSQL = strSQL & " sales_group, shipping_group, account_group, discount_group, "
    strSQL = strSQL & " multicolour, "
    strSQL = strSQL & " colour1, colour1_code, colour1_gmc, colour1_ean, "
    strSQL = strSQL & " colour2, colour2_code, colour2_gmc, colour2_ean, "
    strSQL = strSQL & " colour3, colour3_code, colour3_gmc, colour3_ean, "
    strSQL = strSQL & " colour4, colour4_code, colour4_gmc, colour4_ean, "
    strSQL = strSQL & " colour5, colour5_code, colour5_gmc, colour5_ean, "
    strSQL = strSQL & " lifecycle, lifecycle_date, "
    strSQL = strSQL & " lifecycle_discontinued_item, lifecycle_discontinued_date, "
    strSQL = strSQL & " min_order_qty, order_lot, "
    strSQL = strSQL & " serialised, "
    strSQL = strSQL & " set_item, "
    strSQL = strSQL & " set1, set1_colour, set1_qty, "
    strSQL = strSQL & " set2, set2_colour, set2_qty, "
    strSQL = strSQL & " set3, set3_colour, set3_qty, "
    strSQL = strSQL & " set4, set4_colour, set4_qty, "
    strSQL = strSQL & " set5, set5_colour, set5_qty, "
    strSQL = strSQL & " set6, set6_colour, set6_qty, "
    strSQL = strSQL & " set7, set7_colour, set7_qty, "
    strSQL = strSQL & " set8, set8_colour, set8_qty, "
    strSQL = strSQL & " set9, set9_colour, set9_qty, "
    strSQL = strSQL & " set10,set10_colour,set10_qty,"
    strSQL = strSQL & " gross_weight, nett_weight, "
    strSQL = strSQL & " width, height, depth, volume, "
    strSQL = strSQL & " pack_unit, "
    strSQL = strSQL & " tariff_code, vendor, country_origin, "
    strSQL = strSQL & " ean_code, "
    strSQL = strSQL & " fob, trade, rrp, nz_trade, S03, S04, S08"
    strSQL = strSQL & " mod_required, "
    strSQL = strSQL & " comments, "
    strSQL = strSQL & " created_by "
    strSQL = strSQL & ") VALUES ( "
    strSQL = strSQL & "'" & strDepartment & "',"
    strSQL = strSQL & "'" & Server.HTMLEncode(strBaseCode) & "', '" & Server.HTMLEncode(strItemName) & "', '" & Server.HTMLEncode(strModelName) & "', '" & Server.HTMLEncode(strDescription) & "',"
    strSQL = strSQL & "'" & strGMC & "',"
    strSQL = strSQL & "'" & strSalesGroup & "', '" & strShippingGroup & "', '" & strAccountGroup & "', '" & strDiscountGroup & "',"
    strSQL = strSQL & "'" & intMulticolour & "',"
    strSQL = strSQL & "'" & strColour1 & "', '" & strColour1Code & "', '" & strColour1GMC & "', '" & strColour1EAN & "',"
    strSQL = strSQL & "'" & strColour2 & "', '" & strColour2Code & "', '" & strColour2GMC & "', '" & strColour2EAN & "',"
    strSQL = strSQL & "'" & strColour3 & "', '" & strColour3Code & "', '" & strColour3GMC & "', '" & strColour3EAN & "',"
    strSQL = strSQL & "'" & strColour4 & "', '" & strColour4Code & "', '" & strColour4GMC & "', '" & strColour4EAN & "',"
    strSQL = strSQL & "'" & strColour5 & "', '" & strColour5Code & "', '" & strColour5GMC & "', '" & strColour5EAN & "',"
    strSQL = strSQL & "'" & strLifecycle & "',"
    strSQL = strSQL & " CONVERT(datetime,'" & strLifecycleDate & "',103),"
    strSQL = strSQL & "'" & strLifecycleDiscontinuedItem & "'," 
    strSQL = strSQL & " CONVERT(datetime,'" & strLifecycleDiscontinuedDate & "',103),"
    strSQL = strSQL & "'" & strMinOrderQty & "', '" & strOrderLot & "',"
    strSQL = strSQL & "'" & intSerialised & "',"
    strSQL = strSQL & "'" & intSetItem & "',"
    strSQL = strSQL & "'" & strSet1 & "', '" & strSet1Colour & "', '" & strSet1Qty & "',"
    strSQL = strSQL & "'" & strSet2 & "', '" & strSet2Colour & "', '" & strSet2Qty & "',"
    strSQL = strSQL & "'" & strSet3 & "', '" & strSet3Colour & "', '" & strSet3Qty & "',"
    strSQL = strSQL & "'" & strSet4 & "', '" & strSet4Colour & "', '" & strSet4Qty & "',"
    strSQL = strSQL & "'" & strSet5 & "', '" & strSet5Colour & "', '" & strSet5Qty & "',"
    strSQL = strSQL & "'" & strSet6 & "', '" & strSet6Colour & "', '" & strSet6Qty & "',"
    strSQL = strSQL & "'" & strSet7 & "', '" & strSet7Colour & "', '" & strSet7Qty & "',"
    strSQL = strSQL & "'" & strSet8 & "', '" & strSet8Colour & "', '" & strSet8Qty & "',"
    strSQL = strSQL & "'" & strSet9 & "', '" & strSet9Colour & "', '" & strSet9Qty & "',"
    strSQL = strSQL & "'" & strSet10 & "','" & strSet10Colour & "','" & strSet10Qty & "',"
    strSQL = strSQL & "'" & strGrossWeight & "', '" & strNettWeight & "',"
    strSQL = strSQL & "'" & strWidth & "', '" & strHeight & "', '" & strDepth & "', '" & strVolume & "',"
    strSQL = strSQL & "'" & strPackUnit & "',"
    strSQL = strSQL & "'" & strTariffCode & "', '" & strVendor & "', '" & strCountryOrigin & "',"
    strSQL = strSQL & "'" & strEanCode & "',"
    strSQL = strSQL & "'" & strFOB & "', '" & strTrade & "', '" & strRRP & "', '" & strNzTrade & "','" & strS03 & "','" & strS04 & "','" & strS08 & "',"
    strSQL = strSQL & "'" & intModRequired & "',"
    strSQL = strSQL & "'" & Server.HTMLEncode(strComments) & "',"
    strSQL = strSQL & "'" & session("UsrUserName") & "')"

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
    conn.Execute strSQL  'save item maintenance record

    'MAIN GMC Code Update
    strSQL = "UPDATE tbl_gmc_lookup set ITEM = '" + Server.HTMLEncode(strBaseCode) + "' where GMC = '" + strGMC + "'"
    on error resume next
    conn.Execute strSQL  'update GMC Lookup Table

    if strColour1GMC <> "" then
        strSQL = "UPDATE tbl_gmc_lookup set ITEM = '" + strColour1Code + "' where GMC = '" + strColour1GMC + "'"
        on error resume next
        conn.Execute strSQL  'update GMC Lookup Table
    end if

    if strColour2GMC <> "" then
        strSQL = "UPDATE tbl_gmc_lookup set ITEM = '" + strColour2Code + "' where GMC = '" + strColour2GMC + "'"
        on error resume next
        conn.Execute strSQL  'update GMC Lookup Table
    end if

    if strColour3GMC <> "" then
        strSQL = "UPDATE tbl_gmc_lookup set ITEM = '" + strColour3Code + "' where GMC = '" + strColour3GMC + "'"
        on error resume next
        conn.Execute strSQL  'update GMC Lookup Table
    end if

    if strColour4GMC <> "" then
        strSQL = "UPDATE tbl_gmc_lookup set ITEM = '" + strColour4Code + "' where GMC = '" + strColour4GMC + "'"
        on error resume next
        conn.Execute strSQL  'update GMC Lookup Table
    end if

    if strColour5GMC <> "" then
        strSQL = "UPDATE tbl_gmc_lookup set ITEM = '" + strColour5Code + "' where GMC = '" + strColour5GMC + "'"
        on error resume next
        conn.Execute strSQL  'update GMC Lookup Table
    end if

    'Main EAN Code update
    strSQL = "UPDATE tbl_ean_lookup set ITEM = '" + Server.HTMLEncode(strBaseCode) + "', DATE_ENTERED = getdate() where EAN = '" + strEanCode + "'"
    on error resume next
    conn.Execute strSQL  'update EAN Lookup Table

    if strColour1EAN <> "" then
        strSQL = "UPDATE tbl_ean_lookup set ITEM = '" + strColour1Code + "', DATE_ENTERED = getdate() where EAN = '" + strColour1EAN + "'"
        on error resume next
        conn.Execute strSQL  'update EAN Lookup Table
    end if

    if strColour2EAN <> "" then
        strSQL = "UPDATE tbl_ean_lookup set ITEM = '" + strColour2Code + "', DATE_ENTERED = getdate() where EAN = '" + strColour2EAN + "'"
        on error resume next
        conn.Execute strSQL  'update EAN Lookup Table
    end if

    if strColour3EAN <> "" then
        strSQL = "UPDATE tbl_ean_lookup set ITEM = '" + strColour3Code + "', DATE_ENTERED = getdate() where EAN = '" + strColour3EAN + "'"
        on error resume next
        conn.Execute strSQL  'update EAN Lookup Table
    end if

    if strColour4EAN <> "" then
        strSQL = "UPDATE tbl_ean_lookup set ITEM = '" + strColour4Code + "', DATE_ENTERED = getdate() where EAN = '" + strColour4EAN + "'"
        on error resume next
        conn.Execute strSQL  'update EAN Lookup Table
    end if

    if strColour5EAN <> "" then
        strSQL = "UPDATE tbl_ean_lookup set ITEM = '" + strColour5Code + "', DATE_ENTERED = getdate() where EAN = '" + strColour5EAN + "'"
        on error resume next
        conn.Execute strSQL  'update EAN Lookup Table
    end if

    'On error Goto 0  

    if error <> 0 then
        strMessageText = err.description
    else
        Select Case strDepartment
            case "AV"
                'emailTo    = "russell.wykes@music.yamaha.com"
                emailTo     = "simon.goldsworthy@music.yamaha.com"
            case "MPD"
                emailTo     = "michael.shade@music.yamaha.com"
                emailCc     = "joseph.pantalleresco@music.yamaha.com"
        end select
        emailSubject        = "New Item Maintenance - Approval needed"

        From = "automailer@music.yamaha.com"

        emailBodyText = "G'day!" & vbCrLf _
                      & " " & vbCrLf _
                      & "There is a new item maintenance that requires approval from you." & vbCrLf _
                      & " " & vbCrLf _
                      & "The item: " & strBaseCode & " was requested by " & session("UsrUserName") & vbCrLf _
                      & " " & vbCrLf _
                      & "Please click on the below link to approve it:" & vbCrLf _
                      & "http://intranet/logistics/list_item-maintenance.asp" & vbCrLf _
                      & " " & vbCrLf _
                      & "Thank you. (This is an automated email - please do not reply to this email)"

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

        Response.Redirect("thank-you_item-maintenance.asp")
    end if

    call CloseDataBase()
end sub

sub setGMCCode
    dim intRecordCount

    call OpenDataBase()

    set rs = Server.CreateObject("ADODB.recordset")

    rs.CursorLocation = 3      'adUseClient
    rs.CursorType = 3          'adOpenStatic
    rs.PageSize = 100

    strSQL = "SELECT top 6 GMC, Y6SOSC FROM tbl_gmc_lookup " + _
             "left join as400.s1027cfg.ygzflib.yf6ml01 on y6sos2 = GMC " + _
             "where y6sosc is null and item = '' order by GMC asc "

    rs.Open strSQL, conn

    intRecordCount = rs.recordcount
    For intRecord = 0 To intRecordCount - 1
        strNewGMCCode(intRecord) = rs("GMC")
        rs.movenext
        If rs.EOF Then Exit For
    Next
    'strNewGMCCode = rs("GMC")
    call CloseDataBase()
end sub

sub setEANCode
    call OpenDataBase()

    set rs = Server.CreateObject("ADODB.recordset")

    rs.CursorLocation = 3   'adUseClient
    rs.CursorType = 3       'adOpenStatic
    rs.PageSize = 100

    strSQL = "SELECT top 6 EAN FROM tbl_ean_lookup " + _
             "left join as400.s1027cfg.ygzflib.yf3ml01 on y3eanc = EAN " + _
             "where y3sosc is null and item is null order by EAN asc "

    rs.Open strSQL, conn

    intRecordCount = rs.recordcount
    For intRecord = 0 To intRecordCount - 1
        strNewEANCode(intRecord) = rs("EAN")
        rs.movenext
        If rs.EOF Then Exit For
    Next

    call CloseDataBase()
end sub

sub main
    call UTL_validateLogin 

    call getSalesGroupList
    call getShippingGroupList
    call getAccountGroupList
    call getDiscountGroupList
    call getTariffCodeList
    call getCountryList
    call getLifecycleList


    if Request.ServerVariables("REQUEST_METHOD") = "POST" then
        if Trim(Request("Action")) = "Add" then
            call addItemMaintenance
        end if
    end if

    if Request("ycjitem") =  "NO" then
        call setGMCCode
        call setEANCode
    end if
end sub

dim strSalesGroupList
dim strShippingGroupList
dim strAccountGroupList
dim strDiscountGroupList
dim strTariffCodeList
dim strCountryList
dim strLifecycleList
dim strNewGMCCode(6)
dim strNewEANCode(6)

call main
%>
</head>
<body>

<div class="overlay">
    <div id="loading-spinner"></div>
</div>

<form action="" method="post" name="form_add_item_maintenance" id="form_add_item_maintenance" onsubmit="return validateItemMaintenanceForm(this)">
  <table width="100%" cellpadding="0" cellspacing="0">
    <!-- #include file="include/header.asp" -->
    <tr>
      <td class="first_content">
        <table cellpadding="5" cellspacing="0" border="0">
          <tr>
            <td><a href="list_item-maintenance.asp"><img src="images/icon_item-maintenance.jpg" border="0" alt="Item Maintenance" /></a></td>
            <td valign="top"><img src="images/backward_arrow.gif" border="0" /> <a href="list_item-maintenance.asp">Back to List</a>
              <h2>Add Item Maintenance</h2>
              <font color="green"><%= strMessageText %></font></td>
          </tr>
        </table>
        <h3 style="color:red">Please ensure that you are <u>not</u> submitting the same item twice. Thank you.</h3>
        <table border="0" cellpadding="3" cellspacing="0" class="wide_table">
          <tr>
            <td width="33%" align="left" valign="top"><table cellpadding="3" cellspacing="0" class="item_maintenance_box">
                <tr>
                  <td class="item_maintenance_header" colspan="2">1. Item Details</td>
                </tr>
                <tr>
                  <td width="25%">YCJ Item?<span class="mandatory">*</span>:</td>
                  <td width="75%">
                    <select name="cboYCJItem" onchange="YCJItemSet();">
                      <option value="">...</option>
                      <option <% if Request("cboYCJItem") = "YES" then Response.Write " selected" end if%> value="YES" >YES</option>
                      <option <% if Request("cboYCJItem") = "NO" then Response.Write " selected" end if%> value="NO" >NO</option>
                    </select>
                  </td>
                </tr>
                <tr>
                  <td width="25%">Department<span class="mandatory">*</span>:</td>
                  <td width="75%"><select name="cboDepartment">
                    <option value="">...</option>
                      <option value="AV">AV</option>
                      <option value="MPD">MPD</option>
                    </select></td>
                </tr>
                <tr>
                  <td>BASE code<span class="mandatory">*</span>:</td>
                  <td><input type="text" id="txtBaseCode" name="txtBaseCode" maxlength="15" size="20" />
                    <small>(eg. BRX750B)</small></td>
                </tr>
                <tr>
                  <td>Item name<span class="mandatory">*</span>:</td>
                  <td><input type="text" id="txtItemName" name="txtItemName" maxlength="30" size="30" />
                    <small>(eg. Centre unit for MCR755)</small></td>
                </tr>
                <tr>
                  <td>Model name<span class="mandatory">*</span>:</td>
                  <td><input type="text" id="txtModelName" name="txtModelName" maxlength="20" size="30" />
                    <small>(eg. BRX-750 Black)</small></td>
                </tr>
                <tr>
                  <td>Description<span class="mandatory">*</span>:</td>
                  <td><input type="text" id="txtDescription" name="txtDescription" maxlength="20" size="30" />
                    <small>(eg. BD Receiver Black)</small></td>
                </tr>
                <tr>
                  <td>GMC code<span class="mandatory">*</span>:</td>
                  <td><input type="text" id="txtGMCcode" name="txtGMCcode" maxlength="20" size="30" value="<%= strNewGMCCode(0) %>"/></td>
                </tr>
                <tr>
                  <td>Sales group:</td>
                  <td><select name="cboSalesGroup">
                      <%= strSalesGroupList %>
                    </select></td>
                </tr>
                <tr>
                  <td>Shipping group:</td>
                  <td><select name="cboShippingGroup">
                      <%= strShippingGroupList %>
                    </select></td>
                </tr>
                <tr>
                  <td>Account group:</td>
                  <td><select name="cboAccountGroup">
                      <%= strAccountGroupList %>
                    </select></td>
                </tr>
                <tr>
                  <td>Multicolour?</td>
                  <td><select name="cboMulticolour">
                      <option value="0" rel="none">No</option>
                      <option value="1" rel="multicolour">Yes</option>
                    </select></td>
                </tr>
                <tr rel="multicolour">
                  <td colspan="2"><table cellpadding="0" cellspacing="0" width="100%" border="0">
                      <tr class="documents_all">
                        <td width="30%">BASE code<span class="mandatory">*</span></td>
                        <td width="30%">Description<span class="mandatory">*</span></td>
                        <td width="20%">GMC<span class="mandatory">*</span></td>
                        <td width="15%">EAN</td>
                        <td width="5%">&nbsp;</td>
                      </tr>
                      <tr class="documents_all">
                        <td width="30%" align="left">1:<input type="text" id="txtColour1BaseCode" name="txtColour1BaseCode" maxlength="15" size="15" /></td>
                        <td width="30%" align="left"><input type="text" id="txtColour1" name="txtColour1" maxlength="20" size="20" /></td>
                        <td width="20%" align="left"><input type="text" id="txtGMC1" name="txtGMC1" maxlength="15" size="10" value="<%= strNewGMCCode(1) %>"/></td>
                        <td width="15%" align="left"><input type="text" id="txtEAN1" name="txtEAN1" maxlength="15" size="10" value="<%= strNewEANCode(1) %>"/></td>
                        <td width="5%"><h3 class="myheadings">+color</h3></td>
                      </tr>
                      <tr>
                        <td colspan="5"><div class="thecontent">
                            <table cellpadding="0" cellspacing="0" width="100%" border="0">
                              <tr class="documents_all">
                                <td width="30%" align="left">2:<input type="text" id="txtColour2BaseCode" name="txtColour2BaseCode" maxlength="15" size="15" /></td>
                                <td width="30%" align="left"><input type="text" id="txtColour2" name="txtColour2" maxlength="20" size="20" /></td>
                                <td width="20%" align="left"><input type="text" id="txtGMC2" name="txtGMC2" maxlength="15" size="10" value="<%= strNewGMCCode(2) %>"/></td>
                                <td width="15%" align="left"><input type="text" id="txtEAN2" name="txtEAN2" maxlength="15" size="10" value="<%= strNewEANCode(2) %>"/></td>
                                <td width="5%"><h3 class="myheadings">+color</h3></td>
                              </tr>
                            </table>
                          </div></td>
                      </tr>
                      <tr>
                        <td colspan="5"><div class="thecontent">
                            <table cellpadding="0" cellspacing="0" width="100%" border="0">
                              <tr class="documents_all">
                                <td width="30%" align="left">3:<input type="text" id="txtColour3BaseCode" name="txtColour3BaseCode" maxlength="15" size="15" /></td>
                                <td width="30%" align="left"><input type="text" id="txtColour3" name="txtColour3" maxlength="20" size="20" /></td>
                                <td width="20%" align="left"><input type="text" id="txtGMC3" name="txtGMC3" maxlength="15" size="10" value="<%= strNewGMCCode(3) %>"/></td>
                                <td width="15%" align="left"><input type="text" id="txtEAN3" name="txtEAN3" maxlength="15" size="10" value="<%= strNewEANCode(3) %>"/></td>
                                <td width="5%"><h3 class="myheadings">+color</h3></td>
                              </tr>
                            </table>
                          </div></td>
                      </tr>
                      <tr>
                        <td colspan="5"><div class="thecontent">
                            <table cellpadding="0" cellspacing="0" width="100%" border="0">
                              <tr class="documents_all">
                                <td width="30%" align="left">4:<input type="text" id="txtColour4BaseCode" name="txtColour4BaseCode" maxlength="15" size="15" /></td>
                                <td width="30%" align="left"><input type="text" id="txtColour4" name="txtColour4" maxlength="20" size="20" /></td>
                                <td width="20%" align="left"><input type="text" id="txtGMC4" name="txtGMC4" maxlength="15" size="10" value="<%= strNewGMCCode(4) %>"/></td>
                                <td width="15%" align="left"><input type="text" id="txtEAN4" name="txtEAN4" maxlength="15" size="10" value="<%= strNewEANCode(4) %>"/></td>
                                <td width="5%"><h3 class="myheadings">+color</h3></td>
                              </tr>
                            </table>
                          </div></td>
                      </tr>
                      <tr>
                        <td colspan="5"><div class="thecontent">
                            <table cellpadding="0" cellspacing="0" width="100%" border="0">
                              <tr class="documents_all">
                                <td width="30%" align="left">5:<input type="text" id="txtColour5BaseCode" name="txtColour5BaseCode" maxlength="15" size="15" /></td>
                                <td width="30%" align="left"><input type="text" id="txtColour5" name="txtColour5" maxlength="20" size="20" /></td>
                                <td width="20%" align="left"><input type="text" id="txtGMC5" name="txtGMC5" maxlength="15" size="10" value="<%= strNewGMCCode(5) %>"/></td>
                                <td width="20%" align="left"><input type="text" id="txtEAN5" name="txtEAN5" maxlength="15" size="10" value="<%= strNewEANCode(5) %>"/></td>
                                <td width="5%">&nbsp;</td>
                              </tr>
                            </table>
                          </div></td>
                      </tr>
                    </table></td>
                </tr>
              </table>
              <br />
              <table cellpadding="3" cellspacing="0" class="item_maintenance_box">
                <tr>
                  <td class="item_maintenance_header" colspan="4">2. General</td>
                </tr>
                <tr>
                  <td width="30%">Lifecycle:</td>
                  <td width="30%"><select name="cboLifecycle">
                      <option value="C" rel="none">Current</option>
                      <option value="D" rel="discontinued">D: Discontinued</option>
                      <option value="E" rel="none">E: Not EMC Compliant</option>
                      <option value="H" rel="none">H: Incomplete Item</option>
                      <option value="N" rel="none">N: Stock to be Held</option>
                      <option value="W" rel="none">W: Parts Excessive Stocks</option>
                    </select></td>
                  <td width="20%">&nbsp;</td>
                  <td width="20%">&nbsp;</td>
                </tr>
                <tr>
                  <td>Lifecycle expiry date:</td>
                  <td><input type="text" id="txtLifecycleDate" name="txtLifecycleDate" maxlength="10" size="10" onchange="return validateDate(this)" />
                    <em>DD/MM/YYYY</em></td>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
                <tr rel="discontinued">
                  <td colspan="4"><table width="100%">
                      <tr class="documents_all">
                        <td width="50%">Altern. Item:
                          <input type="text" id="txtAlternativeItem" name="txtAlternativeItem" maxlength="20" size="20" /></td>
                        <td width="50%">Date:
                          <input type="text" id="txtAlternativeItemDate" name="txtAlternativeItemDate" maxlength="10" size="10" />
                          <em>DD/MM/YYYY</em></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td>Min order qty:</td>
                  <td><input type="text" id="txtMinOrderQty" name="txtMinOrderQty" maxlength="5" size="5" /></td>
                  <td>Order lot:</td>
                  <td><input type="text" id="txtOrderLot" name="txtOrderLot" maxlength="5" size="5" /></td>
                </tr>
                <tr>
                  <td colspan="2"><input type="checkbox" name="chkSerialised" id="chkSerialised" value="1" /> Serialised </td>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
              </table></td>
            <td width="33%" valign="top"><table cellpadding="3" cellspacing="0" class="item_maintenance_box">
                <tr>
                  <td class="item_maintenance_header" colspan="2">3. Set Item / Kit Item*</td>
                </tr>
                <tr>
                  <td width="20%">Set / Kit*?</td>
                  <td width="80%"><select name="cboSetItem">
                      <option value="0" rel="none">No</option>
                      <option value="1" rel="setitem">Set</option>
                      <option value="2" rel="setitem">Kit</option>
                    </select> 
                  </td>
                </tr>
                <tr rel="setitem">
                  <td colspan="2"><table cellpadding="0" cellspacing="0" width="100%" border="0">
                      <tr class="documents_all">
                        <td>&nbsp;</td>
                        <td>BASE code<span class="mandatory">*</span></td>
                        <td>Qty<span class="mandatory">*</span></td>
                        <td>&nbsp;</td>
                      </tr>
                      <tr class="documents_all">
                        <td width="30%"># 1:
                          <select name="cboItem1Colour">
                            <option value="">...</option>
                            <option value="1">Colour 1</option>
                            <option value="2">Colour 2</option>
                            <option value="3">Colour 3</option>
                            <option value="4">Colour 4</option>
                            <option value="5">Colour 5</option>
                          </select></td>
                        <td width="30%"><input type="text" id="txtItem1" name="txtItem1" maxlength="15" size="20" /></td>
                        <td width="25%"><input type="text" id="txtQty1" name="txtQty1" maxlength="4" size="5" /></td>
                        <td width="15%" nowrap="nowrap"><h3 class="myheadings">+</h3></td>
                      </tr>
                      <tr>
                        <td colspan="4"><div class="thecontent">
                            <table cellpadding="0" cellspacing="0" width="100%" border="0">
                              <tr class="documents_all">
                                <td width="30%"># 2:
                                  <select name="cboItem2Colour">
                                    <option value="">...</option>
                                    <option value="1">Colour 1</option>
                                    <option value="2">Colour 2</option>
                                    <option value="3">Colour 3</option>
                                    <option value="4">Colour 4</option>
                                    <option value="5">Colour 5</option>
                                  </select></td>
                                <td width="30%"><input type="text" id="txtItem2" name="txtItem2" maxlength="15" size="20" /></td>
                                <td width="25%"><input type="text" id="txtQty2" name="txtQty2" maxlength="4" size="5" /></td>
                                <td width="15%"><h3 class="myheadings">+</h3></td>
                              </tr>
                            </table>
                          </div></td>
                      </tr>
                      <tr>
                        <td colspan="4"><div class="thecontent">
                            <table cellpadding="0" cellspacing="0" width="100%" border="0">
                              <tr class="documents_all">
                                <td width="30%"># 3:
                                  <select name="cboItem3Colour">
                                    <option value="">...</option>
                                    <option value="1">Colour 1</option>
                                    <option value="2">Colour 2</option>
                                    <option value="3">Colour 3</option>
                                    <option value="4">Colour 4</option>
                                    <option value="5">Colour 5</option>
                                  </select></td>
                                <td width="30%"><input type="text" id="txtItem3" name="txtItem3" maxlength="15" size="20" /></td>
                                <td width="25%"><input type="text" id="txtQty3" name="txtQty3" maxlength="4" size="5" /></td>
                                <td width="15%"><h3 class="myheadings">+</h3></td>
                              </tr>
                            </table>
                          </div></td>
                      </tr>
                      <tr>
                        <td colspan="4"><div class="thecontent">
                            <table cellpadding="0" cellspacing="0" width="100%" border="0">
                              <tr class="documents_all">
                                <td width="30%"># 4:
                                  <select name="cboItem4Colour">
                                    <option value="">...</option>
                                    <option value="1">Colour 1</option>
                                    <option value="2">Colour 2</option>
                                    <option value="3">Colour 3</option>
                                    <option value="4">Colour 4</option>
                                    <option value="5">Colour 5</option>
                                  </select></td>
                                <td width="30%"><input type="text" id="txtItem4" name="txtItem4" maxlength="15" size="20" /></td>
                                <td width="25%"><input type="text" id="txtQty4" name="txtQty4" maxlength="4" size="5" /></td>
                                <td width="15%"><h3 class="myheadings">+</h3></td>
                              </tr>
                            </table>
                          </div></td>
                      </tr>
                      <tr>
                        <td colspan="4"><div class="thecontent">
                            <table cellpadding="0" cellspacing="0" width="100%" border="0">
                              <tr class="documents_all">
                                <td width="30%"># 5:
                                  <select name="cboItem5Colour">
                                    <option value="">...</option>
                                    <option value="1">Colour 1</option>
                                    <option value="2">Colour 2</option>
                                    <option value="3">Colour 3</option>
                                    <option value="4">Colour 4</option>
                                    <option value="5">Colour 5</option>
                                  </select></td>
                                <td width="30%"><input type="text" id="txtItem5" name="txtItem5" maxlength="15" size="20" /></td>
                                <td width="25%"><input type="text" id="txtQty5" name="txtQty5" maxlength="4" size="5" /></td>
                                <td width="15%"><h3 class="myheadings">+</h3></td>
                              </tr>
                            </table>
                          </div></td>
                      </tr>
                      <tr>
                        <td colspan="4"><div class="thecontent">
                            <table cellpadding="0" cellspacing="0" width="100%" border="0">
                              <tr class="documents_all">
                                <td width="30%"># 6:
                                  <select name="cboItem6Colour">
                                    <option value="">...</option>
                                    <option value="1">Colour 1</option>
                                    <option value="2">Colour 2</option>
                                    <option value="3">Colour 3</option>
                                    <option value="4">Colour 4</option>
                                    <option value="5">Colour 5</option>
                                  </select></td>
                                <td width="30%"><input type="text" id="txtItem6" name="txtItem6" maxlength="15" size="20" /></td>
                                <td width="25%"><input type="text" id="txtQty6" name="txtQty6" maxlength="4" size="5" /></td>
                                <td width="15%"><h3 class="myheadings">+</h3></td>
                              </tr>
                            </table>
                          </div></td>
                      </tr>
                      <tr>
                        <td colspan="4"><div class="thecontent">
                            <table cellpadding="0" cellspacing="0" width="100%" border="0">
                              <tr class="documents_all">
                                <td width="30%"># 7:
                                  <select name="cboItem7Colour">
                                    <option value="">...</option>
                                    <option value="1">Colour 1</option>
                                    <option value="2">Colour 2</option>
                                    <option value="3">Colour 3</option>
                                    <option value="4">Colour 4</option>
                                    <option value="5">Colour 5</option>
                                  </select></td>
                                <td width="30%"><input type="text" id="txtItem7" name="txtItem7" maxlength="15" size="20" /></td>
                                <td width="25%"><input type="text" id="txtQty7" name="txtQty7" maxlength="4" size="5" /></td>
                                <td width="15%"><h3 class="myheadings">+</h3></td>
                              </tr>
                            </table>
                          </div></td>
                      </tr>
                      <tr>
                        <td colspan="4"><div class="thecontent">
                            <table cellpadding="0" cellspacing="0" width="100%" border="0">
                              <tr class="documents_all">
                                <td width="30%"># 8:
                                  <select name="cboItem8Colour">
                                    <option value="">...</option>
                                    <option value="1">Colour 1</option>
                                    <option value="2">Colour 2</option>
                                    <option value="3">Colour 3</option>
                                    <option value="4">Colour 4</option>
                                    <option value="5">Colour 5</option>
                                  </select></td>
                                <td width="30%"><input type="text" id="txtItem8" name="txtItem8" maxlength="15" size="20" /></td>
                                <td width="25%"><input type="text" id="txtQty8" name="txtQty8" maxlength="4" size="5" /></td>
                                <td width="15%"><h3 class="myheadings">+</h3></td>
                              </tr>
                            </table>
                          </div></td>
                      </tr>
                      <tr>
                        <td colspan="4"><div class="thecontent">
                            <table cellpadding="0" cellspacing="0" width="100%" border="0">
                              <tr class="documents_all">
                                <td width="30%"># 9:
                                  <select name="cboItem9Colour">
                                    <option value="">...</option>
                                    <option value="1">Colour 1</option>
                                    <option value="2">Colour 2</option>
                                    <option value="3">Colour 3</option>
                                    <option value="4">Colour 4</option>
                                    <option value="5">Colour 5</option>
                                  </select></td>
                                <td width="30%"><input type="text" id="txtItem9" name="txtItem9" maxlength="15" size="20" /></td>
                                <td width="25%"><input type="text" id="txtQty9" name="txtQty9" maxlength="4" size="5" /></td>
                                <td width="15%"><h3 class="myheadings">+</h3></td>
                              </tr>
                            </table>
                          </div></td>
                      </tr>
                      <tr>
                        <td colspan="4"><div class="thecontent">
                            <table cellpadding="0" cellspacing="0" width="100%" border="0">
                              <tr class="documents_all">
                                <td width="30%"># 10:
                                  <select name="cboItem10Colour">
                                    <option value="">...</option>
                                    <option value="1">Colour 1</option>
                                    <option value="2">Colour 2</option>
                                    <option value="3">Colour 3</option>
                                    <option value="4">Colour 4</option>
                                    <option value="5">Colour 5</option>
                                  </select></td>
                                <td width="30%"><input type="text" id="txtItem10" name="txtItem10" maxlength="15" size="20" /></td>
                                <td width="25%"><input type="text" id="txtQty10" name="txtQty10" maxlength="4" size="5" /></td>
                                <td width="15%">&nbsp;</td>
                              </tr>
                            </table>
                          </div></td>
                      </tr>
                    </table></td>
                </tr>
              </table>
              <p><strong>Kit:</strong> Created at TT, items to be placed into pre made cartons and all components placed into one outer carton.</p>
                <p><strong>Set:</strong> Multi box items which are picked individually and sent as individual components.</p>
              <table cellpadding="3" cellspacing="0" class="item_maintenance_box">
                <tr>
                  <td class="item_maintenance_header" colspan="4">4. Dimensions</td>
                </tr>
                <tr>
                  <td width="25%">Gross weight<span class="mandatory">*</span>:</td>
                  <td width="25%"><input type="text" id="txtGrossWeight" name="txtGrossWeight" maxlength="5" size="5" />
                    kg</td>
                  <td width="20%">Nett weight<span class="mandatory">*</span>:</td>
                  <td width="30%"><input type="text" id="txtNettWeight" name="txtNettWeight" maxlength="5" size="5" />
                    kg</td>
                </tr>
                <tr>
                  <td>Width<span class="mandatory">*</span>:</td>
                  <td><input type="text" id="txtWidth" name="txtWidth" maxlength="5" size="5" /> cm</td>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
                <tr>
                  <td>Height<span class="mandatory">*</span>:</td>
                  <td><input type="text" id="txtHeight" name="txtHeight" maxlength="5" size="5" /> cm</td>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
                <tr>
                  <td>Depth<span class="mandatory">*</span>:</td>
                  <td><input type="text" id="txtDepth" name="txtDepth" maxlength="5" size="5" /> cm</td>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
                <tr>
                  <td>Packaging qty:</td>
                  <td><input type="text" id="txtPackUnit" name="txtPackUnit" maxlength="5" size="5" /></td>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
              </table>
              <br />
              <table cellpadding="3" cellspacing="0" class="item_maintenance_box">
                <tr>
                  <td class="item_maintenance_header" colspan="2">5. Logistics</td>
                </tr>
                <tr>
                  <td width="30%">Tariff code:</td>
                  <td width="70%"><select name="cboTariffCode">
                      <%= strTariffCodeList %>
                    </select></td>
                </tr>
                <tr>
                  <td>Vendor?</td>
                  <td><select name="cboVendor">
                      <option value="">...</option>
                      <option value="Yamaha">Yamaha</option>
                      <option value="Steinberg">Steinberg</option>
                      <option value="Paiste">Paiste</option>
                      <option value="VOX">VOX</option>
                      <option value="other">Other</option>
                    </select></td>
                </tr>
                <!--<tr rel="othervendor">
                  <td align="right">Please specify:</td>
                  <td><input type="text" id="txtVendorOther" name="txtVendorOther" maxlength="30" size="30" /></td>
                </tr>-->
                <tr>
                  <td>Country of origin:</td>
                  <td><select name="cboCountryOrigin">
                      <%= strCountryList %>
                    </select></td>
                </tr>
                <tr>
                  <td>EAN code<span class="mandatory">*</span>:</td>
                  <td><input type="text" id="txtEANcode" name="txtEANcode" maxlength="30" size="40" value="<%= strNewEANCode(0) %>" /></td>
                </tr>
              </table></td>
            <td width="33%" valign="top"><table cellpadding="3" cellspacing="0" class="item_maintenance_box">
                <tr>
                  <td class="item_maintenance_header" colspan="2">6. Pricing</td>
                </tr>
                <tr>
                  <td>FOB (P01):</td>
                  <td>$
                    <input type="text" id="txtFOB" name="txtFOB" maxlength="10" size="10" /></td>
                </tr>
                <tr>
                  <td width="30%">Trade (S01 ex tax)<span class="mandatory">*</span>:</td>
                  <td width="70%">$
                    <input type="text" id="txtTrade" name="txtTrade" maxlength="10" size="10" /></td>
                </tr>
                <tr>
                  <td>RRP (S50 inc tax)<span class="mandatory">*</span>:</td>
                  <td>$
                    <input type="text" id="txtRRP" name="txtRRP" maxlength="10" size="10" /></td>
                </tr>
                <tr>
                  <td>NZ Trade (S02 ex tax):</td>
                  <td>$
                    <input type="text" id="txtNZtrade" name="txtNZtrade" maxlength="10" size="10" /></td>
                </tr>
				<tr>
                  <td>NZ USD (S03):</td>
                  <td>$
                    <input type="text" id="txtS03" name="txtS03" maxlength="10" size="10" /></td>
                </tr>
				<tr>
                  <td>Narta (S04):</td>
                  <td>$
                    <input type="text" id="txtS04" name="txtS04" maxlength="10" size="10" /></td>
                </tr>
				<tr>
                  <td>AV Custom (S08):</td>
                  <td>$
                    <input type="text" id="txtS08" name="txtS08" maxlength="10" size="10" /></td>
                </tr>
                <tr>
                  <td>Mod required?</td>
                  <td><select name="cboModRequired">
                      <option value="0">No</option>
                      <option value="1">Yes</option>
                    </select></td>
                </tr>
                <tr>
                  <td>Discount group:</td>
                  <td><select name="cboDiscountGroup">
                      <%= strDiscountGroupList %>
                    </select></td>
                </tr>
              </table>
              <br />
              <table cellpadding="3" cellspacing="0" class="item_maintenance_box">
                <tr>
                  <td class="item_maintenance_header">7. Additional Info (Max 200 chars)</td>
                </tr>
                <tr>
                  <td width="70%"><textarea name="txtComments" id="txtComments" cols="55" rows="5" onKeyDown="limitText(this.form.txtComments,this.form.countdown,200);" 
onKeyUp="limitText(this.form.txtComments,this.form.countdown,200);"></textarea></td>
                </tr>
              </table></td>
          </tr>
        </table>
        <p>
          <input type="hidden" name="Action" />
          <input type="submit" value="Add Item Maintenance" />
          <input type="reset" value="Reset" />
        </p></td>
    </tr>
  </table>
</form>
</body>
</html>