<% response.cookies("current_URL_cookie_logistics") = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?" & Request.Querystring %>
<!--#include file="include/connection_it.asp " -->
<!--#include file="class/clsItemMaintenance.asp " -->
<!--#include file="class/clsComment.asp " -->
<% strSection = "item" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<META HTTP-EQUIV="Expires" CONTENT="-1">
<title>Update Item Maintenance</title>
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="include/usableforms.js"></script>
<script type="text/javascript" src="include/generic_form_validations.js"></script>
<script type="text/javascript" src="include/javascript.js"></script>
<script type="text/javascript" src="include/jquery.min.js"></script>
<script language="JavaScript" type="text/javascript">
function validateItemMaintenanceForm(theForm) {
    var reason = "";
    var blnSubmit = true;

    reason += validateEmptyField(theForm.txtBaseCode,"BASE Code");
    reason += validateSpecialCharacters(theForm.txtBaseCode,"BASE Code");
    reason += validateEmptyField(theForm.txtItemName,"Item Name");
    reason += validateSpecialCharacters(theForm.txtItemName,"Item Name");
    reason += validateEmptyField(theForm.txtModelName,"Model Name");
    reason += validateSpecialCharacters(theForm.txtModelName,"Model Name");
    reason += validateEmptyField(theForm.txtDescription,"Description");
    reason += validateSpecialCharacters(theForm.txtDescription,"Description");
    reason += validateEmptyField(theForm.txtDescription,"Description");
    reason += validateSpecialCharacters(theForm.txtGMCcode,"GMC Code");
    reason += validateEmptyField(theForm.txtGrossWeight,"Gross Weight");
    reason += validateSpecialCharacters(theForm.txtGrossWeight,"Gross Weight");
    reason += validateEmptyField(theForm.txtNettWeight,"Nett Weight");
    reason += validateSpecialCharacters(theForm.txtNettWeight,"Nett Weight");
    reason += validateNumeric(theForm.txtWidth,"Width");
    reason += validateNumeric(theForm.txtHeight,"Height");
    reason += validateNumeric(theForm.txtDepth,"Depth");
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
        theForm.Action.value = 'Update';

        return true;
    }
}

function submitMarketingApprove(theForm) {
    var blnSubmit = true;

    if (blnSubmit == true){
        //if (theForm.cboGMapproval.value == 0) {
        //  theForm.Action.value = 'GM did not Approve';
        //} else {
        theForm.Action.value = 'Marketing Approve';
        //}

        return true;
    }
}

function submitGMapprove(theForm) {
    var blnSubmit = true;

    if (blnSubmit == true){
        //if (theForm.cboGMapproval.value == 0) {
        //  theForm.Action.value = 'GM did not Approve';
        //} else {
        theForm.Action.value = 'GM Approve';
        //}

        return true;
    }
}

function submitEMCapprove(theForm) {
    var blnSubmit = true;

    if (blnSubmit == true) {
        theForm.Action.value = 'EMC Approved';

        return true;
    }
}

function submitLogisticsProcessed(theForm) {
    var blnSubmit = true;

    if (blnSubmit == true) {
        theForm.Action.value = 'Logistics Processed';

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
'-----------------------------------------------
' GET ITEM MAINTENANCE
'-----------------------------------------------
Sub getItemMaintenance
    dim strSQL
    dim intID
    intID = request("id")

    call OpenDataBase()

    set rs = Server.CreateObject("ADODB.recordset")

    rs.CursorLocation = 3   'adUseClient
    rs.CursorType = 3       'adOpenStatic

    strSQL = "SELECT yma_item_maintenance.*, tbl_users.email, tbl_users.division "
    strSQL = strSQL & " FROM yma_item_maintenance "
    strSQL = strSQL & "     INNER JOIN tbl_users ON yma_item_maintenance.created_by = tbl_users.username "
    strSQL = strSQL & " WHERE item_id = " & intID
    'Response.Write strSQL
    rs.Open strSQL, conn

    if not DB_RecSetIsEmpty(rs) Then
        session("department")                   = rs("department")
        session("base_code")                    = rs("base_code")
        session("item_name")                    = rs("item_name")
        session("model_name")                   = rs("model_name")
        session("description")                  = rs("description")
        session("gmc_code")                     = rs("gmc_code")
        session("sales_group")                  = rs("sales_group")
        session("shipping_group")               = rs("shipping_group")
        session("account_group")                = rs("account_group")
        session("discount_group")               = rs("discount_group")
        session("multicolour")                  = rs("multicolour")
        session("colour1")                      = rs("colour1")
        session("colour1_code")                 = rs("colour1_code")
        session("colour1_gmc")                  = rs("colour1_gmc")
        session("colour1_ean")                  = rs("colour1_ean")
        session("colour2")                      = rs("colour2")
        session("colour2_code")                 = rs("colour2_code")
        session("colour2_gmc")                  = rs("colour2_gmc")
        session("colour2_ean")                  = rs("colour2_ean")
        session("colour3")                      = rs("colour3")
        session("colour3_code")                 = rs("colour3_code")
        session("colour3_gmc")                  = rs("colour3_gmc")
        session("colour3_ean")                  = rs("colour3_ean")
        session("colour4")                      = rs("colour4")
        session("colour4_code")                 = rs("colour4_code")
        session("colour4_gmc")                  = rs("colour4_gmc")
        session("colour4_ean")                  = rs("colour4_ean")
        session("colour5")                      = rs("colour5")
        session("colour5_code")                 = rs("colour5_code")
        session("colour5_gmc")                  = rs("colour5_gmc")
        session("colour5_ean")                  = rs("colour5_ean")
        session("lifecycle")                    = rs("lifecycle")
        session("lifecycle_date")               = rs("lifecycle_date")
        session("lifecycle_discontinued_item")  = rs("lifecycle_discontinued_item")
        session("lifecycle_discontinued_date")  = rs("lifecycle_discontinued_date")
        session("min_order_qty")                = rs("min_order_qty")
        session("order_lot")                    = rs("order_lot")
        session("serialised")                   = rs("serialised")
        session("set_item")                     = rs("set_item")
        session("set1")                         = rs("set1")
        session("set1_colour")                  = rs("set1_colour")
        session("set1_qty")                     = rs("set1_qty")
        session("set2")                         = rs("set2")
        session("set2_colour")                  = rs("set2_colour")
        session("set2_qty")                     = rs("set2_qty")
        session("set3")                         = rs("set3")
        session("set3_colour")                  = rs("set3_colour")
        session("set3_qty")                     = rs("set3_qty")
        session("set4")                         = rs("set4")
        session("set4_colour")                  = rs("set4_colour")
        session("set4_qty")                     = rs("set4_qty")
        session("set5")                         = rs("set5")
        session("set5_colour")                  = rs("set5_colour")
        session("set5_qty")                     = rs("set5_qty")
        session("set6")                         = rs("set6")
        session("set6_colour")                  = rs("set6_colour")
        session("set6_qty")                     = rs("set6_qty")
        session("set7")                         = rs("set7")
        session("set7_colour")                  = rs("set7_colour")
        session("set7_qty")                     = rs("set7_qty")
        session("set8")                         = rs("set8")
        session("set8_colour")                  = rs("set8_colour")
        session("set8_qty")                     = rs("set8_qty")
        session("set9")                         = rs("set9")
        session("set9_colour")                  = rs("set9_colour")
        session("set9_qty")                     = rs("set9_qty")
        session("set10")                        = rs("set10")
        session("set10_colour")                 = rs("set10_colour")
        session("set10_qty")                    = rs("set10_qty")
        session("gross_weight")                 = rs("gross_weight")
        session("nett_weight")                  = rs("nett_weight")
        session("width")                        = rs("width")
        session("height")                       = rs("height")
        session("depth")                        = rs("depth")
        'session("volume")                       = session("width") * session("height") * session("depth")
        session("volume")                       = rs("volume")
        session("item_size")                    = rs("item_size")
        session("pack_unit")                    = rs("pack_unit")
        session("tariff_code")                  = rs("tariff_code")
        session("vendor")                       = rs("vendor")
        session("country_origin")               = rs("country_origin")
        session("ean_code")                     = rs("ean_code")
        session("fob")                          = rs("fob")
        session("trade")                        = rs("trade")
        session("rrp")                          = rs("rrp")
        session("nz_trade")                     = rs("nz_trade")
        session("mod_required")                 = rs("mod_required")
        session("comments")                     = rs("comments")
        session("status")                       = rs("status")
        session("marketing_approval")           = rs("marketing_approval")
        session("marketing_approval_date")      = rs("marketing_approval_date")
        session("gm_approval")                  = rs("gm_approval")
        session("gm_approval_date")             = rs("gm_approval_date")
        session("emc_approval")                 = rs("emc_approval")
        session("emc_approval_date")            = rs("emc_approval_date")
        session("logistics_processed")          = rs("logistics_processed")
        session("logistics_processed_date")     = rs("logistics_processed_date")
        session("created_by")                   = rs("created_by")
        session("date_created")                 = rs("date_created")
        session("modified_by")                  = rs("modified_by")
        session("date_modified")                = rs("date_modified")
        session("requester_email")              = rs("email")
        session("requester_division")           = rs("division")
        session("logistics_pending")            = rs("logistics_pending")
    end if

    call CloseDataBase()
end sub

'-----------------------------------------------
' UPDATE ITEM MAINTENANCE table
'-----------------------------------------------
sub updateItemMaintenance
    dim strSQL
    dim intID
    intID = request("id")

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
    'strVolume                       = Replace(Request.Form("txtVolume"),"'","''")
    'strItemSize                     = Replace(Request.Form("txtItemSize"),"'","''")
    strPackUnit                     = Replace(Request.Form("txtPackUnit"),"'","''")
    strTariffCode                   = Request.Form("cboTariffCode")
    strVendor                       = Replace(Request.Form("txtVendor"),"'","''")
    strCountryOrigin                = Request.Form("cboCountryOrigin")
    strEanCode                      = Replace(Request.Form("txtEANcode"),"'","''")
    strFOB                          = Replace(Request.Form("txtFOB"),"'","''")
    strTrade                        = Replace(Request.Form("txtTrade"),"'","''")
    strRRP                          = Replace(Request.Form("txtRRP"),"'","''")
    strNzTrade                      = Replace(Request.Form("txtNZtrade"),"'","''")
    intModRequired                  = Request.Form("cboModRequired")
    strComments                     = Replace(Request.Form("txtComments"),"'","''")
    intGMapproval                   = Request.Form("cboGMapproval")
    intEMCapproval                  = Request.Form("cboEMCapproval")
    intLogisticsProcessed           = Request.Form("cboLogisticsProcessed")
    intStatus                       = Request.Form("cboStatus")

    Call OpenDataBase()

    strSQL = "UPDATE yma_item_maintenance SET "
    strSQL = strSQL & "department = '" & strDepartment & "',"
    strSQL = strSQL & "base_code = '" & strBaseCode & "',"
    strSQL = strSQL & "item_name = '" & strItemName & "',"
    strSQL = strSQL & "model_name = '" & strModelName & "',"
    strSQL = strSQL & "description = '" & strDescription & "',"
    strSQL = strSQL & "gmc_code = '" & strGMC & "',"
    strSQL = strSQL & "sales_group = '" & strSalesGroup & "',"
    strSQL = strSQL & "shipping_group = '" & strShippingGroup & "',"
    strSQL = strSQL & "account_group = '" & strAccountGroup & "',"
    strSQL = strSQL & "discount_group = '" & strDiscountGroup & "',"
    strSQL = strSQL & "multicolour = '" & intMultiColour & "',"
    strSQL = strSQL & "colour1 = '" & strColour1 & "',"
    strSQL = strSQL & "colour1_code = '" & strColour1Code & "',"
    strSQL = strSQL & "colour1_gmc = '" & strColour1GMC & "',"
    strSQL = strSQL & "colour1_ean = '" & strColour1EAN & "',"
    strSQL = strSQL & "colour2 = '" & strColour2 & "',"
    strSQL = strSQL & "colour2_code = '" & strColour2Code & "',"
    strSQL = strSQL & "colour2_gmc = '" & strColour2GMC & "',"
    strSQL = strSQL & "colour2_ean = '" & strColour2EAN & "',"
    strSQL = strSQL & "colour3 = '" & strColour3 & "',"
    strSQL = strSQL & "colour3_code = '" & strColour3Code & "',"
    strSQL = strSQL & "colour3_gmc = '" & strColour3GMC & "',"
    strSQL = strSQL & "colour3_ean = '" & strColour3EAN & "',"
    strSQL = strSQL & "colour4 = '" & strColour4 & "',"
    strSQL = strSQL & "colour4_code = '" & strColour4Code & "',"
    strSQL = strSQL & "colour4_gmc = '" & strColour4GMC & "',"
    strSQL = strSQL & "colour4_ean = '" & strColour4EAN & "',"
    strSQL = strSQL & "colour5 = '" & strColour5 & "',"
    strSQL = strSQL & "colour5_code = '" & strColour5Code & "',"
    strSQL = strSQL & "colour5_gmc = '" & strColour5GMC & "',"
    strSQL = strSQL & "colour5_ean = '" & strColour5EAN & "',"
    strSQL = strSQL & "lifecycle = '" & strLifecycle & "',"
    strSQL = strSQL & "lifecycle_date = CONVERT(datetime,'" & strLifecycleDate & "',103),"
    strSQL = strSQL & "lifecycle_discontinued_item = '" & strLifecycleDiscontinuedItem & "',"
    strSQL = strSQL & "lifecycle_discontinued_date = CONVERT(datetime,'" & strLifecycleDiscontinuedDate & "',103),"
    strSQL = strSQL & "min_order_qty = '" & strMinOrderQty & "',"
    strSQL = strSQL & "order_lot = '" & strOrderLot & "',"
    strSQL = strSQL & "serialised = '" & intSerialised & "',"
    strSQL = strSQL & "set_item = '" & intSetItem & "',"
    strSQL = strSQL & "set1 = '" & strSet1 & "',"
    strSQL = strSQL & "set1_colour = '" & strSet1Colour & "',"
    strSQL = strSQL & "set1_qty = '" & strSet1Qty & "',"
    strSQL = strSQL & "set2 = '" & strSet2 & "',"
    strSQL = strSQL & "set2_colour = '" & strSet2Colour & "',"
    strSQL = strSQL & "set2_qty = '" & strSet2Qty & "',"
    strSQL = strSQL & "set3 = '" & strSet3 & "',"
    strSQL = strSQL & "set3_colour = '" & strSet3Colour & "',"
    strSQL = strSQL & "set3_qty = '" & strSet3Qty & "',"
    strSQL = strSQL & "set4 = '" & strSet4 & "',"
    strSQL = strSQL & "set4_colour = '" & strSet4Colour & "',"
    strSQL = strSQL & "set4_qty = '" & strSet4Qty & "',"
    strSQL = strSQL & "set5 = '" & strSet5 & "',"
    strSQL = strSQL & "set5_colour = '" & strSet5Colour & "',"
    strSQL = strSQL & "set5_qty = '" & strSet5Qty & "',"
    strSQL = strSQL & "set6 = '" & strSet6 & "',"
    strSQL = strSQL & "set6_colour = '" & strSet6Colour & "',"
    strSQL = strSQL & "set6_qty = '" & strSet6Qty & "',"
    strSQL = strSQL & "set7 = '" & strSet7 & "',"
    strSQL = strSQL & "set7_colour = '" & strSet7Colour & "',"
    strSQL = strSQL & "set7_qty = '" & strSet7Qty & "',"
    strSQL = strSQL & "set8 = '" & strSet8 & "',"
    strSQL = strSQL & "set8_colour = '" & strSet8Colour & "',"
    strSQL = strSQL & "set8_qty = '" & strSet8Qty & "',"
    strSQL = strSQL & "set9 = '" & strSet9 & "',"
    strSQL = strSQL & "set9_colour = '" & strSet9Colour & "',"
    strSQL = strSQL & "set9_qty = '" & strSet9Qty & "',"
    strSQL = strSQL & "set10 = '" & strSet10 & "',"
    strSQL = strSQL & "set10_colour = '" & strSet10Colour & "',"
    strSQL = strSQL & "set10_qty = '" & strSet10Qty & "',"
    strSQL = strSQL & "gross_weight = '" & strGrossWeight & "',"
    strSQL = strSQL & "nett_weight = '" & strNettWeight & "',"
    strSQL = strSQL & "width = '" & strWidth & "',"
    strSQL = strSQL & "height = '" & strHeight & "',"
    strSQL = strSQL & "depth = '" & strDepth & "',"
    strSQL = strSQL & "volume = '" & strVolume & "',"
    'strSQL = strSQL & "item_size = '" & strItemSize & "',"
    strSQL = strSQL & "pack_unit = '" & strPackUnit & "',"
    strSQL = strSQL & "tariff_code = '" & strTariffCode & "',"
    strSQL = strSQL & "vendor = '" & strVendor & "',"
    strSQL = strSQL & "country_origin = '" & strCountryOrigin & "',"
    strSQL = strSQL & "ean_code = '" & strEANcode & "',"
    strSQL = strSQL & "fob = '" & strFOB & "',"
    strSQL = strSQL & "trade = '" & strTrade & "',"
    strSQL = strSQL & "rrp = '" & strRRP & "',"
    strSQL = strSQL & "nz_trade = '" & strNzTrade & "',"
    strSQL = strSQL & "mod_required = '" & intModRequired & "',"
    strSQL = strSQL & "comments = '" & strComments & "',"
    strSQL = strSQL & "date_modified = getdate(),"
    strSQL = strSQL & "modified_by = '" & session("UsrUserName") & "',"
    strSQL = strSQL & "status = '" & intStatus & "' WHERE item_id = " & intID

    'response.Write strSQL
    on error resume next
    conn.Execute strSQL

    if err <> 0 then
        strMessageText = err.description
    else
        strMessageText = "The Item Maintenance has been updated."
    end if

    Call CloseDataBase()
end sub

'-----------------------------------------------
' ITEM LOGS: List all Item Logs
'-----------------------------------------------
sub displayItemLogs
    dim strSQL
    dim intID
    intID = request("id")

    dim strPageResultNumber
    dim strRecordPerPage
    dim intRecordCount

    call OpenDataBase()

    set rs = Server.CreateObject("ADODB.recordset")

    rs.CursorLocation = 3   'adUseClient
    rs.CursorType = 3       'adOpenStatic

    strPageResultNumber = trim(Request("cboDealerResultSize"))
    strRecordPerPage = 200

    rs.PageSize = 200

    strSQL = "SELECT * FROM tbl_item_logs WHERE item_id = " & intID & " ORDER BY log_date"

    rs.Open strSQL, conn

    'Response.Write strSQL

    intPageCount = rs.PageCount
    intRecordCount = rs.recordcount

    strLogList = ""

    if not DB_RecSetIsEmpty(rs) Then
        For intRecord = 1 To rs.PageSize

            strLogList = strLogList & "<tr class=""documents_all"">"
            strLogList = strLogList & "<td>" & rs("username") & "</td>"
            strLogList = strLogList & "<td>" & rs("activity") & "</td>"
            strLogList = strLogList & "<td>" & rs("log_date") & "</td>"
            strLogList = strLogList & "</tr>"

            rs.movenext

            If rs.EOF Then Exit For
        next
    else
        strLogList = "<tr><td colspan=3 align=center>&nbsp;</td></tr>"
    end if

    strLogList = strLogList & "<tr>"

    Set rs = nothing
    call CloseDataBase()
end sub

'-----------------------------------------------
' ITEM LOGS: Add record to Item Log table
'-----------------------------------------------
sub addItemLog
    dim strSQL
    dim strAction
    strAction = Trim(Request("Action"))

    dim intID
    intID = request("id")

    Call OpenDataBase()

    strSQL = "INSERT INTO tbl_item_logs (item_id, username, activity, log_date) VALUES ("
    strSQL = strSQL & "'" & intID & "',"
    strSQL = strSQL & "'" & session("UsrUserName") & "',"
    strSQL = strSQL & "'" & strAction & "',"
    strSQL = strSQL & "getdate())"

    'response.Write strSQL
    on error resume next
    conn.Execute strSQL

    if err <> 0 then
        strMessageText = err.description
    else
        'strMessageText = "<br>Item log is added."
    end if

    Call CloseDatabase()
end sub

'-----------------------------------------------
' GM APPROVAL: List GM Approval in IM Approval table
'-----------------------------------------------
sub displayGMapprovalList
    dim strSQL
    dim intID
    intID = request("id")

    dim strPageResultNumber
    dim strRecordPerPage
    dim intRecordCount

    call OpenDataBase()

    set rs = Server.CreateObject("ADODB.recordset")

    rs.CursorLocation = 3   'adUseClient
    rs.CursorType = 3       'adOpenStatic

    strPageResultNumber = trim(Request("cboDealerResultSize"))
    strRecordPerPage = 200

    rs.PageSize = 200

    strSQL = "SELECT * FROM yma_item_maintenance_approval WHERE item_id = " & intID & " ORDER BY approval_date"

    rs.Open strSQL, conn

    'Response.Write strSQL

    intPageCount = rs.PageCount
    intRecordCount = rs.recordcount

    strGMapprovalList = ""

    if not DB_RecSetIsEmpty(rs) Then
        'rs.AbsolutePage = session("cinitialPage")

        For intRecord = 1 To rs.PageSize 

            strGMapprovalList = strGMapprovalList & "<tr>"
            strGMapprovalList = strGMapprovalList & "<td>" & rs("username") & "</td>"
            strGMapprovalList = strGMapprovalList & "<td>" & rs("activity") & "</td>"
            strGMapprovalList = strGMapprovalList & "<td>" & rs("log_date") & "</td>"
            strGMapprovalList = strGMapprovalList & "</tr>"

            rs.movenext

            If rs.EOF Then Exit For
        next
    else
        strGMapprovalList = "<tr><td colspan=3 align=center>&nbsp;</td></tr>"
    end if

    strGMapprovalList = strGMapprovalList & "<tr>"

    call CloseDataBase()
end sub

'-----------------------------------------------
' GM APPROVAL: Add record to IM Approval table
'-----------------------------------------------
sub addGMapproval
    dim strSQL
    dim strAction
    strAction = Trim(Request("Action"))

    dim intID
    intID = request("id")

    dim intApproved
    intApproved = request("cboGMapproval")

    dim strComments
    strComments = Replace(Request.Form("txtGMapprovalComments"),"'","''")

    Call OpenDataBase()

    strSQL = "INSERT INTO yma_item_maintenance_approval (item_id, approval_type, approved, comments, approved_by, approval_date) VALUES ("
    strSQL = strSQL & "'" & intID & "',"
    strSQL = strSQL & "'" & strAction & "',"
    strSQL = strSQL & "'" & intApproved & "',"
    strSQL = strSQL & "'" & strComments & "',"
    strSQL = strSQL & "'" & session("UsrUserName") & "',"
    strSQL = strSQL & "getdate())"

    'response.Write strSQL
    on error resume next
    conn.Execute strSQL

    if err <> 0 then
        strMessageText = err.description
    else
        strMessageText = "<br>Item Log (GM Approval) is added."
    end if

    Call CloseDataBase()
end sub

'-----------------------------------------------
' 2a. If AV Marketing approves, send email to get GM Approval
'-----------------------------------------------
sub requestGMapprovalEmail
    dim intID
    intID = request("id")

    dim strMarketingComments
    strMarketingComments = Replace(Request.Form("txtMarketingApprovalComments"),"'","''")

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

    emailFrom   = "noreply@music.yamaha.com"
    emailTo     = "simon.goldsworthy@music.yamaha.com"
    emailCc     = "russell.wykes@music.yamaha.com"

    emailSubject = "New Item Maintenance - GM Approval needed"

    emailBodyText = "Hi Simon," & vbCrLf _
                  & " " & vbCrLf _
                  & "There is a new item maintenance that requires approval from you." & vbCrLf _
                  & " " & vbCrLf _
                  & "The item (" & session("base_code") & " - created by " & session("created_by") & ") has been approved by: " & session("UsrUserName") & vbCrLf _
                  & " " & vbCrLf _
                  & "Comments: " & strMarketingComments & vbCrLf _
                  & " " & vbCrLf _
                  & "Please click on the below link to approve it: (your username: simong and your default password: password)" & vbCrLf _
                  & "http://intranet/logistics/update_item-maintenance.asp?id=" & intID & "" & vbCrLf _
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
end sub

'------------------------------------------------------
' 2b. If Marketing not approves, send email back to the requester
'------------------------------------------------------
sub sendMarketingNotApprovedEmail
    dim intID
    intID = request("id")

    dim strMarketingComments
    strMarketingComments = Replace(Request.Form("txtMarketingApprovalComments"),"'","''")

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

    emailFrom   = "noreply@music.yamaha.com"
    emailTo     = session("requester_email")
    emailCc     = "russell.wykes@music.yamaha.com"

    emailSubject = "New Item Maintenance - not approved by Marketing"

    emailBodyText = "Hi there," & vbCrLf _
                  & " " & vbCrLf _
                  & "The item (" & session("base_code") & ") that you submitted was NOT approved by: " & session("UsrUserName") & vbCrLf _
                  & " " & vbCrLf _
                  & "Comments: " & strMarketingComments & vbCrLf _
                  & " " & vbCrLf _
                  & "Please click on the below link to update the item:" & vbCrLf _
                  & "http://intranet/logistics/update_item-maintenance.asp?id=" & intID & "" & vbCrLf _
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
end sub

'-----------------------------------------------
' 3a. If GM approves, send email to get EMC Approval
'-----------------------------------------------
sub requestEmcApprovalEmail
    dim intID
    intID = request("id")

    dim strGMComments
    strGMComments = Replace(Request.Form("txtGMapprovalComments"),"'","''")

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

    emailFrom   = "noreply@music.yamaha.com"
    emailTo     = "drew.morrow@music.yamaha.com"
    'emailCc     = "matthew.madden@music.yamaha.com"
    emailBcc    = "logistics-aus@music.yamaha.com"

    emailSubject = "New Item Maintenance - EMC Approval needed"

    emailBodyText = "Hi Drew," & vbCrLf _
                  & " " & vbCrLf _
                  & "There is a new item maintenance that requires EMC approval from you." & vbCrLf _
                  & " " & vbCrLf _
                  & "The item (" & session("base_code") & " - created by " & session("created_by") & ") has been approved by: " & session("UsrUserName") & vbCrLf _
                  & " " & vbCrLf _
                  & "Comments: " & strGMComments & vbCrLf _
                  & " " & vbCrLf _
                  & "Please click on the below link to approve it:" & vbCrLf _
                  & "http://intranet/logistics/update_item-maintenance.asp?id=" & intID & "" & vbCrLf _
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
end sub

'------------------------------------------------------
' 3b. If GM not approves, send email back to the requester
'------------------------------------------------------
sub sendGMnotApprovedEmail
    dim intID
    intID = request("id")

    dim strGMComments
    strGMComments = Replace(Request.Form("txtGMapprovalComments"),"'","''")

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

    emailFrom   = "noreply@music.yamaha.com"
    emailTo     = session("requester_email")
    emailBcc    = "logistics-aus@music.yamaha.com"

    emailSubject = "New Item Maintenance - not approved by GM"

    emailBodyText = "Hi there," & vbCrLf _
                  & " " & vbCrLf _
                  & "The item (" & session("base_code") & ") that you submitted was NOT approved by: " & session("UsrUserName") & vbCrLf _
                  & " " & vbCrLf _
                  & "Comments: " & strGMComments & vbCrLf _
                  & " " & vbCrLf _
                  & "Please click on the below link to update the item:" & vbCrLf _
                  & "http://intranet/logistics/update_item-maintenance.asp?id=" & intID & "" & vbCrLf _
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
end sub

'------------------------------------------------------
' 4a. If EMC is approved, then proceeds to Logistics
'------------------------------------------------------
sub requestLogisticsProcessedEmail
    dim intID
    intID = request("id")

    dim strEMCComments
    strEMCComments = Replace(Request.Form("txtEMCapprovalComments"),"'","''")

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

    emailFrom   = "noreply@music.yamaha.com"
    emailTo     = "logistics-aus@music.yamaha.com"

    emailSubject = "New Item Maintenance - needs to be processed"

    emailBodyText = "Hello Logistics," & vbCrLf _
                  & " " & vbCrLf _
                  & "There is a new item (" & session("base_code") & " created by " & session("created_by") & ") that needs to be processed." & vbCrLf _
                  & " " & vbCrLf _
                  & "Comments: " & strEMCComments & vbCrLf _
                  & " " & vbCrLf _
                  & "Please click on the below link to view the item:" & vbCrLf _
                  & "http://intranet/logistics/update_item-maintenance.asp?id=" & intID & "" & vbCrLf _
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
end sub

'------------------------------------------------------
' 4b. If EMC not approves, send email back to the requester
'------------------------------------------------------
sub sendEMCnotApprovedEmail
    dim intID
    intID = request("id")

    dim strEMCComments
    strEMCComments = Replace(Request.Form("txtEMCapprovalComments"),"'","''")

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

    emailFrom   = "noreply@music.yamaha.com"
    emailTo     = session("requester_email")
    emailBcc    = "logistics-aus@music.yamaha.com"

    emailSubject = "New Item Maintenance - not EMC approved by Logistics"

    emailBodyText = "Hi there," & vbCrLf _
                  & " " & vbCrLf _
                  & "The item (" & session("base_code") & ") that you submitted was NOT EMC approved by: " & session("UsrUserName") & vbCrLf _
                  & " " & vbCrLf _
                  & "Comments: " & strEMCComments & vbCrLf _
                  & " " & vbCrLf _
                  & "Please click on the below link to update the item:" & vbCrLf _
                  & "http://intranet/logistics/update_item-maintenance.asp?id=" & intID & "" & vbCrLf _
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
end sub

'----------------------------------------------------------------------
' 5. Notify requester that the item has been processed by Logistics
'----------------------------------------------------------------------
sub sendLogisticsProcessedEmail
    dim intID
    intID = request("id")

    dim strGMComments
    strGMComments = Replace(Request.Form("txtGMapprovalComments"),"'","''")

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

    emailFrom   = "noreply@music.yamaha.com"
    emailTo     = session("requester_email")
    'emailBcc    = "logistics-aus@music.yamaha.com"

    emailSubject = "Your Item Maintenance has been finalised / processed by Logistics"
    emailFrom = "noreply@music.yamaha.com"

    emailBodyText = "Hi there," & vbCrLf _
                  & " " & vbCrLf _
                  & "The item (" & session("base_code") & ") that you submitted has been finalised / processed by Logistics." & vbCrLf _
                  & " " & vbCrLf _
                  & "Please click on the below link to view the item:" & vbCrLf _
                  & "http://intranet/logistics/update_item-maintenance.asp?id=" & intID & "" & vbCrLf _
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
end sub

'-----------------------------------------------
' II. Marketing APPROVAL: Update IM table
'-----------------------------------------------
sub updateItemMaintenanceMarketingApprove
    dim intID
    intID = request("id")

    dim strSQL

    dim intApproved
    intApproved = request("cboMarketingApproval")

    Call OpenDataBase()

    strSQL = "UPDATE yma_item_maintenance SET "
    strSQL = strSQL & "marketing_approval = " & intApproved & ", "
    strSQL = strSQL & "modified_by = '" & session("UsrUserName") & "', "
    strSQL = strSQL & "marketing_approval_date =  getdate(), date_modified = getdate() WHERE item_id = " & intID

    on error resume next
    conn.Execute strSQL

    if err <> 0 then
        strMessageText = err.description
    else
        strMessageText = "Item Maintenance has been approved by Marketing."

        if intApproved = 1 then
            call requestGMapprovalEmail
        else
            call sendMarketingNotApprovedEmail
        end if
    end if

    Call CloseDataBase()
end sub

'-----------------------------------------------
' Marketing APPROVAL: Add record to IM Approval table
'-----------------------------------------------
sub addMarketingApproval
    dim strSQL
    dim strAction
    strAction = Trim(Request("Action"))

    dim intID
    intID = request("id")

    dim intApproved
    intApproved = request("cboMarketingApproval")

    dim strComments
    strComments = Replace(Request.Form("txtMarketingApprovalComments"),"'","''")

    Call OpenDataBase()

    strSQL = "INSERT INTO yma_item_maintenance_approval (item_id, approval_type, approved, comments, approved_by, approval_date) VALUES ("
    strSQL = strSQL & "'" & intID & "',"
    strSQL = strSQL & "'" & strAction & "',"
    strSQL = strSQL & "'" & intApproved & "',"
    strSQL = strSQL & "'" & strComments & "',"
    strSQL = strSQL & "'" & session("UsrUserName") & "',"
    strSQL = strSQL & "getdate())"

    on error resume next
    conn.Execute strSQL

    if err <> 0 then
        strMessageText = err.description
    else
        strMessageText = "<br>Item Log (Marketing approval) is added."
    end if

    Call CloseDataBase()
end sub

'-----------------------------------------------
' III. GM APPROVAL: Update IM table
'-----------------------------------------------
sub updateItemMaintenanceGMapprove
    dim intID
    intID = request("id")

    dim strSQL

    dim intApproved
    intApproved = request("cboGMapproval")

    Call OpenDataBase()

    strSQL = "UPDATE yma_item_maintenance SET "
    strSQL = strSQL & "gm_approval = " & intApproved & ", "
    strSQL = strSQL & "modified_by = '" & session("UsrUserName") & "', "
    strSQL = strSQL & "gm_approval_date =  getdate(), date_modified = getdate() WHERE item_id = " & intID

    'response.Write strSQL
    on error resume next
    conn.Execute strSQL

    if err <> 0 then
        strMessageText = err.description
    else
        strMessageText = "Item Maintenance has been approved by GM."

        if intApproved = 1 then
            call requestEmcApprovalEmail
        else
            call sendGMnotApprovedEmail
        end if
    end if

    Call CloseDataBase()
end sub

'-----------------------------------------------
' EMC APPROVAL: Add record to IM Approval table
'-----------------------------------------------
sub addEMCapproval
    dim strSQL
    dim strAction
    strAction = Trim(Request("Action"))

    dim intID
    intID = request("id")

    dim intApproved
    intApproved = request("cboEMCapproval")

    dim strComments
    strComments = Replace(Request.Form("txtEMCapprovalComments"),"'","''")

    Call OpenDataBase()

    strSQL = "INSERT INTO yma_item_maintenance_approval (item_id, approval_type, approved, comments, approved_by, approval_date) VALUES ("
    strSQL = strSQL & "'" & intID & "',"
    strSQL = strSQL & "'" & strAction & "',"
    strSQL = strSQL & "'" & intApproved & "',"
    strSQL = strSQL & "'" & strComments & "',"
    strSQL = strSQL & "'" & session("UsrUserName") & "',"
    strSQL = strSQL & "getdate())"

    'response.Write strSQL
    on error resume next
    conn.Execute strSQL

    if err <> 0 then
        strMessageText = err.description
    else
        strMessageText = "<br>Item Log (EMC approval) is added."
    end if

    Call CloseDataBase()
end sub

'-----------------------------------------------
' EMC APPROVAL: Update IM table
'-----------------------------------------------
sub updateItemMaintenanceEMCapprove
    dim strSQL
    dim intID
    intID = request("id")

    dim intApproved
    intApproved = request("cboEMCapproval")

    Call OpenDataBase()

    strSQL = "UPDATE yma_item_maintenance SET "
    strSQL = strSQL & "emc_approval = " & intApproved & ", "
    strSQL = strSQL & "modified_by = '" & session("UsrUserName") & "', "
    strSQL = strSQL & "emc_approval_date =  getdate(), date_modified = getdate() WHERE item_id = " & intID

    'response.Write strSQL
    on error resume next
    conn.Execute strSQL

    if err <> 0 then
        strMessageText = err.description
    else
        strMessageText = "Item Maintenance has been EMC approved."

        if intApproved = 1 then
            call requestLogisticsProcessedEmail
        else
            call sendEMCnotApprovedEmail
        end if
    end if

    Call CloseDataBase()
end sub

'-----------------------------------------------
' LOGISTICS PROCESSED APPROVAL: Update IM table
'-----------------------------------------------
sub updateItemMaintenanceLogisticsProcessed
    dim strSQL
    dim intID
    intID = request("id")

    Call OpenDataBase()

    strSQL = "UPDATE yma_item_maintenance SET "
    strSQL = strSQL & "logistics_processed = 1, "
    strSQL = strSQL & "modified_by = '" & session("UsrUserName") & "', "
    strSQL = strSQL & "logistics_processed_date =  getdate(), date_modified = getdate(), status = 0 WHERE item_id = " & intID

    'response.Write strSQL
    on error resume next
    conn.Execute strSQL

    if err <> 0 then
        strMessageText = err.description
    else
        strMessageText = "Item Maintenance has been processed by Logistics."
    end if

    Call CloseDataBase()
end sub

sub main
    call UTL_validateLogin

    dim intID
    intID = request("id")

    if Request.ServerVariables("REQUEST_METHOD") = "POST" then
        select case Trim(Request("Action"))
            case "Update"
                call updateItemMaintenance
                call addItemLog
            case "Marketing Approve"
                call updateItemMaintenanceMarketingApprove
                call addMarketingApproval
                call addItemLog
            case "GM Approve"
                call updateItemMaintenanceGMapprove
                call addGMapproval
                call addItemLog
            case "EMC Approved"
                call updateItemMaintenanceEMCapprove
                call addEMCapproval
                call addItemLog
            case "Logistics Processed"
                call updateItemMaintenanceLogisticsProcessed
                call addItemLog
                call sendLogisticsProcessedEmail
            case "Comment"
                call addComment(intID,itemMaintenanceModuleID)
                call listComments(intID,itemMaintenanceModuleID)
        end select
    end if

    call listComments(intID,itemMaintenanceModuleID)
    call getItemMaintenance
    call getSalesGroupList
    call getShippingGroupList
    call getAccountGroupList
    call getDiscountGroupList
    call getLifecycleList
    call displayItemLogs
    call getTariffCodeList
    call getCountryList
end sub

dim strLogList
dim strMessageText
dim strSalesGroupList
dim strShippingGroupList
dim strAccountGroupList
dim strDiscountGroupList
dim strTariffCodeList
dim strCountryList
dim strLifecycleList
dim strCommentsList
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
dim strVolume
dim strItemSize
dim strPackUnit
dim strTariffCode
dim strVendor
dim strCountryOrigin
dim strEanCode
dim strTrade
dim strRRP
dim strNzTrade
dim intModRequired
dim strComments
dim intGMapproval
dim intEMCapproval
dim intLogisticsProcessed
dim intStatus

call main
%>
</head>
<body>
<input id="UsrLoginRole" type="hidden" name="UsrLoginRole" value="<% Response.Write(session("UsrLoginRole")) %>">
<table cellspacing="0" cellpadding="0" align="center" class="full_size_table" border="0">
  <!-- #include file="include/header.asp" -->
  <tr>
    <td class="first_content"><table cellpadding="5" cellspacing="0" border="0">
        <tr>
          <td><a href="list_item-maintenance.asp"><img src="images/icon_item-maintenance.jpg" border="0" alt="Item Maintenance" /></a></td>
          <td valign="top"><img src="images/backward_arrow.gif" border="0" /> <a href="list_item-maintenance.asp">Back to List</a>
            <h2>Update Item Maintenance</h2>
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
      <div align="right"><img src="images/icon_printer.gif" border="0" /> <a href="javascript:PrintThisPage()">Printable version</a></div>
      <table cellpadding="5" cellspacing="0" class="item_maintenance_approval_box" border="0">
        <tr>
          <td colspan="2" bgcolor="#f0f0f0"><strong>APPROVALS</strong></td>
        </tr>
        <tr>
          <td width="15%" valign="top">GM <img src="images/forward_arrow.gif" border="0" /></td>
          <td width="85%"><% select case session("gm_approval")
                        case "1" %>
            <strong style="color:green">APPROVED</strong> - <%= displayDateFormatted(session("gm_approval_date")) %>
            <%  case "0" %>
            <strong style="color:red">REJECTED</strong> - <%= displayDateFormatted(session("gm_approval_date")) %>
            <table width="600" border="0" cellspacing="0" cellpadding="3" bgcolor="#FF6600">
              <form action="" method="post" onsubmit="return submitGMapprove(this)">
                <tr>
                  <td width="50%">
                    <%' TODO Steven Vranch requested access whilst Mark Amory away. Please revert change in 1 month (Victor Samson 2016-08-19) %>
                    <textarea name="txtGMapprovalComments" id="txtGMapprovalComments" cols="30" rows="2" <% if session("UsrUserName") <> "stevenv" or session("UsrLoginRole") <> 11 then Response.Write("disabled") end if %>><%= session("gm_approval_comments") %></textarea>
                  </td>
                  <td width="50%" valign="top">
                    <%' TODO Steven Vranch requested access whilst Mark Amory away. Please revert change in 1 month (Victor Samson 2016-08-19) %>
                    <select name="cboGMapproval" <% if session("UsrUserName") <> "stevenv" or session("UsrLoginRole") <> 11 then Response.Write("disabled") end if %>>
                      <option <% if session("gm_approval") = "1" then Response.Write " selected" end if%> value="1">Approved</option>
                      <option <% if session("gm_approval") = "0" then Response.Write " selected" end if%> value="0">Not approved</option>
                    </select>
                    <input type="hidden" name="Action" />
                    <%' TODO Steven Vranch requested access whilst Mark Amory away. Please revert change in 1 month (Victor Samson 2016-08-19) %>
                    <input type="submit" value="GM Approve" <% if session("UsrUserName") <> "stevenv" or session("UsrLoginRole") <> 11 then Response.Write("disabled") end if %> /></td>
                </tr>
              </form>
            </table>
            <%  case else %>
            <strong style="color:orange">WAITING FOR APPROVAL</strong>
            <table width="600" border="0" cellspacing="0" cellpadding="3" bgcolor="#00CC66">
              <form action="" method="post" onsubmit="return submitGMapprove(this)">
                <tr>
                  <td width="50%">
                    <%' TODO Steven Vranch requested access whilst Mark Amory away. Please revert change in 1 month (Victor Samson 2016-08-19) %>
                    <textarea name="txtGMapprovalComments" id="txtGMapprovalComments" cols="30" rows="2" <% if session("UsrUserName") <> "stevenv" and session("UsrLoginRole") <> 11 then Response.Write("disabled") end if %>><%= session("gm_approval_comments") %></textarea>
                  </td>
                  <td width="50%" valign="top">
                    <%' TODO Steven Vranch requested access whilst Mark Amory away. Please revert change in 1 month (Victor Samson 2016-08-19) %>
                    <select name="cboGMapproval" <% if session("UsrUserName") <> "stevenv" and session("UsrLoginRole") <> 11 then Response.Write("disabled") end if %>>
                      <option <% if session("gm_approval") = "1" then Response.Write " selected" end if%> value="1">Approved</option>
                      <option <% if session("gm_approval") = "0" then Response.Write " selected" end if%> value="0">Not approved</option>
                    </select>
                    <input type="hidden" name="Action" />
                    <%' TODO Steven Vranch requested access whilst Mark Amory away. Please revert change in 1 month (Victor Samson 2016-08-19) %>
                    <input type="submit" value="GM Approve" <% if session("UsrUserName") <> "stevenv" and session("UsrLoginRole") <> 11 then Response.Write("disabled") end if %> /></td>
                </tr>
              </form>
            </table>
            <% end select %></td>
        </tr>
        <% if session("gm_approval") = "1" then %>
        <tr>
          <td valign="top">EMC <img src="images/forward_arrow.gif" border="0" /></td>
          <td><% select case session("emc_approval") 
                        case "1" %>
            <strong style="color:green">APPROVED</strong> - <%= displayDateFormatted(session("emc_approval_date")) %>
            <%  case "0" %>
            <strong style="color:red">REJECTED</strong> - <%= displayDateFormatted(session("emc_approval_date")) %>
            <table width="600" border="0" cellspacing="0" cellpadding="3" bgcolor="#FF6600">
              <form action="" method="post" onsubmit="return submitEMCapprove(this)">
                <tr>
                  <td width="50%"><textarea name="txtEMCapprovalComments" id="txtEMCapprovalComments" cols="30" rows="2" <% if session("UsrLoginRole") <> 1 then Response.Write("disabled") end if %>><%= session("emc_approval_comments") %></textarea></td>
                  <td width="50%" valign="top"><select name="cboEMCapproval" <% if session("UsrLoginRole") <> 1 then Response.Write("disabled") end if %>>
                      <option <% if session("emc_approval") = "1" then Response.Write " selected" end if%> value="1">Approved</option>
                      <option <% if session("emc_approval") = "0" then Response.Write " selected" end if%> value="0">Not approved</option>
                    </select>
                    <input type="hidden" name="Action" />
                    <input type="submit" value="EMC Approved" <% if session("UsrUserName") <> "drewm" and session("UsrUserName") <> "harsonos" and session("UsrUserName") <> "matthewm" then Response.Write("disabled") end if %> /></td>
                </tr>
              </form>
            </table>
            <%  case else %>
            <strong style="color:orange">WAITING FOR APPROVAL</strong>
            <table width="600" border="0" cellspacing="0" cellpadding="3" bgcolor="#00CC66">
              <form action="" method="post" onsubmit="return submitEMCapprove(this)">
                <tr>
                  <td width="50%"><textarea name="txtEMCapprovalComments" id="txtEMCapprovalComments" cols="30" rows="2" <% if session("UsrUserName") <> "drewm" and session("UsrUserName") <> "harsonos" and session("UsrUserName") <> "matthewm" then Response.Write("disabled") end if %>><%= session("emc_approval_comments") %></textarea></td>
                  <td width="50%" valign="top"><select name="cboEMCapproval" <% if session("UsrUserName") <> "drewm" and session("UsrUserName") <> "harsonos" and session("UsrUserName") <> "matthewm" then Response.Write("disabled") end if %>>
                      <option <% if session("emc_approval") = "1" then Response.Write " selected" end if%> value="1">Approved</option>
                      <option <% if session("emc_approval") = "0" then Response.Write " selected" end if%> value="0">Not approved</option>
                    </select>
                    <input type="hidden" name="Action" />
                    <input type="submit" value="EMC Approved" <% if session("UsrUserName") <> "drewm" and session("UsrUserName") <> "harsonos" and session("UsrUserName") <> "matthewm" then Response.Write("disabled") end if %> /></td>
                </tr>
              </form>
            </table>
            <% end select %></td>
        </tr>
        <% end if %>
        <% if session("gm_approval") = "1" and session("emc_approval") = "1" then %>
        <tr>
          <td valign="top">Logistics <img src="images/forward_arrow.gif" border="0" /></td>
          <td><% if session("logistics_processed") = "1" then %>
            <strong style="color:green">PROCESSED</strong> - <%= displayDateFormatted(session("logistics_processed_date")) %>
            <% else %>
            <strong style="color:orange">WAITING FOR APPROVAL</strong>
            <form action="" method="post" onsubmit="return submitLogisticsProcessed(this)">
              <input type="hidden" name="Action" />
              <input type="submit" value="Logistics Processed" <% if session("UsrLoginRole") <> 1 and session("UsrLoginRole") <> 6 then Response.Write("disabled") end if %> />
            </form>
            <% end if %></td>
        </tr>
        <% end if %>
      </table>
      <br />
      <div id="contentstart">
        <form action="" method="post" name="form_update_item_maintenance" id="form_update_item_maintenance" onsubmit="return validateItemMaintenanceForm(this)">
          <table border="0" cellpadding="3" cellspacing="0" class="wide_table">
            <tr>
              <td width="33%" align="left" valign="top"><table cellpadding="3" cellspacing="0" class="item_maintenance_box" id="item_details">
                  <tr>
                    <td class="item_maintenance_header" colspan="2">1. Item Details</td>
                  </tr>
                  <tr>
                    <td width="30%">Department:</td>
                    <td width="70%"><select name="cboDepartment">
                        <option <% if session("department") = "AV" then Response.Write " selected" end if%> value="AV">AV</option>
                        <option <% if session("department") = "MPD" then Response.Write " selected" end if%> value="MPD">MPD</option>
                      </select></td>
                  </tr>
                  <tr>
                    <td width="30%">BASE code<span class="mandatory">*</span>:</td>
                    <td width="70%"><input type="text" id="txtBaseCode" name="txtBaseCode" maxlength="15" size="20" value="<%= Server.HTMLEncode(session("base_code")) %>" /></td>
                  </tr>
                  <tr>
                    <td>Item name<span class="mandatory">*</span>:</td>
                    <td><input type="text" id="txtItemName" name="txtItemName" maxlength="30" size="35" value="<%= Server.HTMLEncode(session("item_name")) %>" /></td>
                  </tr>
                  <tr>
                    <td>Model name<span class="mandatory">*</span>:</td>
                    <td><input type="text" id="txtModelName" name="txtModelName" maxlength="20" size="30" value="<%= Server.HTMLEncode(session("model_name")) %>" /></td>
                  </tr>
                  <tr>
                    <td>Description<span class="mandatory">*</span>:</td>
                    <td><input type="text" id="txtDescription" name="txtDescription" maxlength="30" size="30" value="<%= Server.HTMLEncode(session("description")) %>" /></td>
                  </tr>
                  <tr>
                    <td>GMC code:</td>
                    <td><input type="text" id="txtGMCcode" name="txtGMCcode" maxlength="20" size="30" value="<%= Server.HTMLEncode(session("gmc_code")) %>" /></td>
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
                        <option <% if session("multicolour") = "0" then Response.Write " selected" end if%> value="0" rel="none">No</option>
                        <option <% if session("multicolour") = "1" then Response.Write " selected" end if%> value="1" rel="multicolour">Yes</option>
                      </select></td>
                  </tr>
                  <tr rel="multicolour">
                    <td colspan="2"><table width="100%">
                        <tr class="documents_all">
                          <td bgcolor="#FFFFFF">&nbsp;</td>
                          <td>BASE code<span class="mandatory">*</span></td>
                          <td>Description<span class="mandatory">*</span></td>
                          <td>GMC<span class="mandatory">*</span></td>
                          <td>EAN:</td>
                        </tr>
                        <tr class="documents_all">
                          <td><small>C1</small></td>
                          <td><input type="text" id="txtColour1BaseCode" name="txtColour1BaseCode" maxlength="15" size="20" value="<%= session("colour1_code") %>" /></td>
                          <td><input type="text" id="txtColour1" name="txtColour1" maxlength="20" size="20" value="<%= Server.HTMLEncode(session("colour1")) %>" /></td>
                          <td><input type="text" id="txtGMC1" name="txtGMC1" maxlength="15" size="10" value="<%= session("colour1_gmc") %>" /></td>
                          <td><input type="text" id="txtEAN1" name="txtEAN1" maxlength="15" size="15" value="<%= session("colour1_ean") %>" /></td>
                        </tr>
                        <tr class="documents_all">
                          <td><small>C2</small></td>
                          <td><input type="text" id="txtColour2BaseCode" name="txtColour2BaseCode" maxlength="15" size="20" value="<%= session("colour2_code") %>" /></td>
                          <td><input type="text" id="txtColour2" name="txtColour2" maxlength="20" size="20" value="<%= Server.HTMLEncode(session("colour2")) %>" /></td>
                          <td><input type="text" id="txtGMC2" name="txtGMC2" maxlength="15" size="10" value="<%= session("colour2_gmc") %>" /></td>
                          <td><input type="text" id="txtEAN2" name="txtEAN2" maxlength="15" size="15" value="<%= session("colour2_ean") %>" /></td>
                        </tr>
                        <tr class="documents_all">
                          <td><small>C3</small></td>
                          <td><input type="text" id="txtColour3BaseCode" name="txtColour3BaseCode" maxlength="15" size="20" value="<%= session("colour3_code") %>" /></td>
                          <td><input type="text" id="txtColour3" name="txtColour3" maxlength="20" size="20" value="<%= Server.HTMLEncode(session("colour3")) %>" /></td>
                          <td><input type="text" id="txtGMC3" name="txtGMC3" maxlength="15" size="10" value="<%= session("colour3_gmc") %>" /></td>
                          <td><input type="text" id="txtEAN3" name="txtEAN3" maxlength="15" size="15" value="<%= session("colour3_ean") %>" /></td>
                        </tr>
                        <tr class="documents_all">
                          <td><small>C4</small></td>
                          <td><input type="text" id="txtColour4BaseCode" name="txtColour4BaseCode" maxlength="15" size="20" value="<%= session("colour4_code") %>" /></td>
                          <td><input type="text" id="txtColour4" name="txtColour4" maxlength="20" size="20" value="<%= Server.HTMLEncode(session("colour4")) %>" /></td>
                          <td><input type="text" id="txtGMC4" name="txtGMC4" maxlength="15" size="10" value="<%= session("colour4_gmc") %>" /></td>
                          <td><input type="text" id="txtEAN4" name="txtEAN4" maxlength="15" size="15" value="<%= session("colour4_ean") %>" /></td>
                        </tr>
                        <tr class="documents_all">
                          <td><small>C5</small></td>
                          <td><input type="text" id="txtColour5BaseCode" name="txtColour5BaseCode" maxlength="15" size="20" value="<%= session("colour5_code") %>" /></td>
                          <td><input type="text" id="txtColour5" name="txtColour5" maxlength="20" size="20" value="<%= Server.HTMLEncode(session("colour5")) %>" /></td>
                          <td><input type="text" id="txtGMC5" name="txtGMC5" maxlength="15" size="10" value="<%= session("colour5_gmc") %>" /></td>
                          <td><input type="text" id="txtEAN5" name="txtEAN5" maxlength="15" size="15" value="<%= session("colour5_ean") %>" /></td>
                        </tr>
                      </table></td>
                  </tr>
                </table>
                <br />
                <table cellpadding="3" cellspacing="0" class="item_maintenance_box" id="general">
                  <tr>
                    <td class="item_maintenance_header" colspan="4">2. General</td>
                  </tr>
                  <tr>
                    <td width="30%">Lifecycle:</td>
                    <td width="30%"><select name="cboLifecycle">
                        <option <% if session("lifecycle") = "C" then Response.Write " selected" end if%> value="C" rel="none">Current</option>
                        <option <% if session("lifecycle") = "D" then Response.Write " selected" end if%> value="D" rel="discontinued">D: Discontinued</option>
                        <option <% if session("lifecycle") = "E" then Response.Write " selected" end if%> value="E" rel="none">E: Not EMC Compliant</option>
                        <option <% if session("lifecycle") = "H" then Response.Write " selected" end if%> value="H" rel="none">H: Incomplete Item</option>
                        <option <% if session("lifecycle") = "N" then Response.Write " selected" end if%> value="N" rel="none">N: Stock to be Held</option>
                        <option <% if session("lifecycle") = "W" then Response.Write " selected" end if%> value="W" rel="none">W: Parts Excessive Stocks</option>
                      </select></td>
                    <td width="15%">&nbsp;</td>
                    <td width="25%">&nbsp;</td>
                  </tr>
                  <tr>
                    <td>Lifecycle expiry date:</td>
                    <td><input type="text" id="txtLifecycleDate" name="txtLifecycleDate" maxlength="10" size="10" value="<%= session("lifecycle_date") %>" />
                      <em>DD/MM/YYYY</em></td>
                    <td>&nbsp;</td>
                    <td>&nbsp;</td>
                  </tr>
                  <tr rel="discontinued">
                    <td colspan="4"><table width="100%">
                        <tr class="documents_all">
                          <td width="50%">Altern. Item:
                            <input type="text" id="txtAlternativeItem" name="txtAlternativeItem" maxlength="20" size="20" value="<%= session("lifecycle_discontinued_item") %>" /></td>
                          <td width="50%">Date:
                            <input type="text" id="txtAlternativeItemDate" name="txtAlternativeItemDate" maxlength="10" size="10" value="<%= session("lifecycle_discontinued_date") %>" />
                            <em>(DD/MM/YYYY)</em></td>
                        </tr>
                      </table></td>
                  </tr>
                  <tr>
                    <td>Min order qty:</td>
                    <td><input type="text" id="txtMinOrderQty" name="txtMinOrderQty" maxlength="5" size="5" value="<%= session("min_order_qty") %>" /></td>
                    <td>Order Lot:</td>
                    <td><input type="text" id="txtOrderLot" name="txtOrderLot" maxlength="5" size="5" value="<%= session("order_lot") %>" /></td>
                  </tr>
                  <tr>
                    <td colspan="2"><input type="checkbox" name="chkSerialised" id="chkSerialised" value="1" <% if session("serialised") = "1" then Response.Write " checked" end if%> /> Serialised </td>
                    <td>&nbsp;</td>
                    <td>&nbsp;</td>
                  </tr>
                </table>
                <br />
                <input type="hidden" name="Action" />
                <input type="submit" value="Update Item Maintenance" <% if session("status") = "0" and session("UsrUserName") <> "drewm" and session("UsrUserName") <> "harsonos" and session("UsrUserName") <> "matthewm" then Response.Write "disabled" end if%> />

                </td>

              <td width="33%" valign="top"><table cellpadding="3" cellspacing="0" class="item_maintenance_box" id="set_item_kit_item">
                  <tr>
                    <td class="item_maintenance_header" colspan="2">3. Set Item / Kit Item*</td>
                  </tr>
                  <tr>
                    <td width="20%">Set / Kit*?</td>
                    <td width="80%"><select name="cboSetItem">
                        <option <% if session("set_item") = "0" then Response.Write " selected" end if%> value="0" rel="none">No</option>
                        <option <% if session("set_item") = "1" then Response.Write " selected" end if%> value="1" rel="setitem">Set</option>
                        <option <% if session("set_item") = "2" then Response.Write " selected" end if%> value="2" rel="setitem">Kit</option>
                      </select></td>
                  </tr>
                  <tr rel="setitem">
                    <td colspan="2"><table width="100%">
                        <tr class="documents_all">
                          <td bgcolor="#FFFFFF">&nbsp;</td>
                          <td>Colour</td>
                          <td>BASE code<span class="mandatory">*</span></td>
                          <td>Qty<span class="mandatory">*</span></td>
                        </tr>
                        <tr class="documents_all">
                          <td><small># 1:</small></td>
                          <td><select name="cboItem1Colour">
                              <option <% if session("set1_colour") = "" then Response.Write " selected" end if%> value="">...</option>
                              <option <% if session("set1_colour") = "1" then Response.Write " selected" end if%> value="1">Colour 1</option>
                              <option <% if session("set1_colour") = "2" then Response.Write " selected" end if%> value="2">Colour 2</option>
                              <option <% if session("set1_colour") = "3" then Response.Write " selected" end if%> value="3">Colour 3</option>
                              <option <% if session("set1_colour") = "4" then Response.Write " selected" end if%> value="4">Colour 4</option>
                              <option <% if session("set1_colour") = "5" then Response.Write " selected" end if%> value="5">Colour 5</option>
                            </select></td>
                          <td><input type="text" id="txtItem1" name="txtItem1" maxlength="15" size="20" value="<%= session("set1") %>" /></td>
                          <td><input type="text" id="txtQty1" name="txtQty1" maxlength="4" size="5" value="<%= session("set1_qty") %>" /></td>
                        </tr>
                        <tr class="documents_all">
                          <td><small># 2:</small></td>
                          <td><select name="cboItem2Colour">
                              <option <% if session("set2_colour") = "" then Response.Write " selected" end if%> value="">...</option>
                              <option <% if session("set2_colour") = "1" then Response.Write " selected" end if%> value="1">Colour 1</option>
                              <option <% if session("set2_colour") = "2" then Response.Write " selected" end if%> value="2">Colour 2</option>
                              <option <% if session("set2_colour") = "3" then Response.Write " selected" end if%> value="3">Colour 3</option>
                              <option <% if session("set2_colour") = "4" then Response.Write " selected" end if%> value="4">Colour 4</option>
                              <option <% if session("set2_colour") = "5" then Response.Write " selected" end if%> value="5">Colour 5</option>
                            </select></td>
                          <td><input type="text" id="txtItem2" name="txtItem2" maxlength="15" size="20" value="<%= session("set2") %>" /></td>
                          <td><input type="text" id="txtQty2" name="txtQty2" maxlength="4" size="5" value="<%= session("set2_qty") %>" /></td>
                        </tr>
                        <tr class="documents_all">
                          <td><small># 3:</small></td>
                          <td><select name="cboItem3Colour">
                              <option <% if session("set3_colour") = "" then Response.Write " selected" end if%> value="">...</option>
                              <option <% if session("set3_colour") = "1" then Response.Write " selected" end if%> value="1">Colour 1</option>
                              <option <% if session("set3_colour") = "2" then Response.Write " selected" end if%> value="2">Colour 2</option>
                              <option <% if session("set3_colour") = "3" then Response.Write " selected" end if%> value="3">Colour 3</option>
                              <option <% if session("set3_colour") = "4" then Response.Write " selected" end if%> value="4">Colour 4</option>
                              <option <% if session("set3_colour") = "5" then Response.Write " selected" end if%> value="5">Colour 5</option>
                            </select></td>
                          <td><input type="text" id="txtItem3" name="txtItem3" maxlength="15" size="20" value="<%= session("set3") %>" /></td>
                          <td><input type="text" id="txtQty3" name="txtQty3" maxlength="4" size="5" value="<%= session("set3_qty") %>" /></td>
                        </tr>
                        <tr class="documents_all">
                          <td><small># 4:</small></td>
                          <td><select name="cboItem4Colour">
                              <option <% if session("set4_colour") = "" then Response.Write " selected" end if%> value="">...</option>
                              <option <% if session("set4_colour") = "1" then Response.Write " selected" end if%> value="1">Colour 1</option>
                              <option <% if session("set4_colour") = "2" then Response.Write " selected" end if%> value="2">Colour 2</option>
                              <option <% if session("set4_colour") = "3" then Response.Write " selected" end if%> value="3">Colour 3</option>
                              <option <% if session("set4_colour") = "4" then Response.Write " selected" end if%> value="4">Colour 4</option>
                              <option <% if session("set4_colour") = "5" then Response.Write " selected" end if%> value="5">Colour 5</option>
                            </select></td>
                          <td><input type="text" id="txtItem4" name="txtItem4" maxlength="15" size="20" value="<%= session("set4") %>" /></td>
                          <td><input type="text" id="txtQty4" name="txtQty4" maxlength="4" size="5" value="<%= session("set4_qty") %>" /></td>
                        </tr>
                        <tr class="documents_all">
                          <td><small># 5:</small></td>
                          <td height="26"><select name="cboItem5Colour">
                              <option <% if session("set5_colour") = "" then Response.Write " selected" end if%> value="">...</option>
                              <option <% if session("set5_colour") = "1" then Response.Write " selected" end if%> value="1">Colour 1</option>
                              <option <% if session("set5_colour") = "2" then Response.Write " selected" end if%> value="2">Colour 2</option>
                              <option <% if session("set5_colour") = "3" then Response.Write " selected" end if%> value="3">Colour 3</option>
                              <option <% if session("set5_colour") = "4" then Response.Write " selected" end if%> value="4">Colour 4</option>
                              <option <% if session("set5_colour") = "5" then Response.Write " selected" end if%> value="5">Colour 5</option>
                            </select></td>
                          <td><input type="text" id="txtItem5" name="txtItem5" maxlength="15" size="20" value="<%= session("set5") %>" /></td>
                          <td><input type="text" id="txtQty5" name="txtQty5" maxlength="4" size="5" value="<%= session("set5_qty") %>" /></td>
                        </tr>
                        <tr class="documents_all">
                          <td><small># 6:</small></td>
                          <td><select name="cboItem6Colour">
                              <option <% if session("set6_colour") = "" then Response.Write " selected" end if%> value="">...</option>
                              <option <% if session("set6_colour") = "1" then Response.Write " selected" end if%> value="1">Colour 1</option>
                              <option <% if session("set6_colour") = "2" then Response.Write " selected" end if%> value="2">Colour 2</option>
                              <option <% if session("set6_colour") = "3" then Response.Write " selected" end if%> value="3">Colour 3</option>
                              <option <% if session("set6_colour") = "4" then Response.Write " selected" end if%> value="4">Colour 4</option>
                              <option <% if session("set6_colour") = "5" then Response.Write " selected" end if%> value="5">Colour 5</option>
                            </select></td>
                          <td><input type="text" id="txtItem6" name="txtItem6" maxlength="15" size="20" value="<%= session("set6") %>" /></td>
                          <td><input type="text" id="txtQty6" name="txtQty6" maxlength="4" size="5" value="<%= session("set6_qty") %>" /></td>
                        </tr>
                        <tr class="documents_all">
                          <td><small># 7:</small></td>
                          <td><select name="cboItem7Colour">
                              <option <% if session("set7_colour") = "" then Response.Write " selected" end if%> value="">...</option>
                              <option <% if session("set7_colour") = "1" then Response.Write " selected" end if%> value="1">Colour 1</option>
                              <option <% if session("set7_colour") = "2" then Response.Write " selected" end if%> value="2">Colour 2</option>
                              <option <% if session("set7_colour") = "3" then Response.Write " selected" end if%> value="3">Colour 3</option>
                              <option <% if session("set7_colour") = "4" then Response.Write " selected" end if%> value="4">Colour 4</option>
                              <option <% if session("set7_colour") = "5" then Response.Write " selected" end if%> value="5">Colour 5</option>
                            </select></td>
                          <td><input type="text" id="txtItem7" name="txtItem7" maxlength="15" size="20" value="<%= session("set7") %>" /></td>
                          <td><input type="text" id="txtQty7" name="txtQty7" maxlength="4" size="5" value="<%= session("set7_qty") %>" /></td>
                        </tr>
                        <tr class="documents_all">
                          <td><small># 8:</small></td>
                          <td><select name="cboItem8Colour">
                              <option <% if session("set8_colour") = "" then Response.Write " selected" end if%> value="">...</option>
                              <option <% if session("set8_colour") = "1" then Response.Write " selected" end if%> value="1">Colour 1</option>
                              <option <% if session("set8_colour") = "2" then Response.Write " selected" end if%> value="2">Colour 2</option>
                              <option <% if session("set8_colour") = "3" then Response.Write " selected" end if%> value="3">Colour 3</option>
                              <option <% if session("set8_colour") = "4" then Response.Write " selected" end if%> value="4">Colour 4</option>
                              <option <% if session("set8_colour") = "5" then Response.Write " selected" end if%> value="5">Colour 5</option>
                            </select></td>
                          <td><input type="text" id="txtItem8" name="txtItem8" maxlength="15" size="20" value="<%= session("set8") %>" /></td>
                          <td><input type="text" id="txtQty8" name="txtQty8" maxlength="4" size="5" value="<%= session("set8_qty") %>" /></td>
                        </tr>
                        <tr class="documents_all">
                          <td><small># 9:</small></td>
                          <td><select name="cboItem9Colour">
                              <option <% if session("set9_colour") = "" then Response.Write " selected" end if%> value="">...</option>
                              <option <% if session("set9_colour") = "1" then Response.Write " selected" end if%> value="1">Colour 1</option>
                              <option <% if session("set9_colour") = "2" then Response.Write " selected" end if%> value="2">Colour 2</option>
                              <option <% if session("set9_colour") = "3" then Response.Write " selected" end if%> value="3">Colour 3</option>
                              <option <% if session("set9_colour") = "4" then Response.Write " selected" end if%> value="4">Colour 4</option>
                              <option <% if session("set9_colour") = "5" then Response.Write " selected" end if%> value="5">Colour 5</option>
                            </select></td>
                          <td><input type="text" id="txtItem9" name="txtItem9" maxlength="15" size="20" value="<%= session("set9") %>" /></td>
                          <td><input type="text" id="txtQty9" name="txtQty9" maxlength="4" size="5" value="<%= session("set9_qty") %>" /></td>
                        </tr>
                        <tr class="documents_all">
                          <td><small># 10:</small></td>
                          <td><select name="cboItem10Colour">
                              <option <% if session("set10_colour") = "" then Response.Write " selected" end if%> value="">...</option>
                              <option <% if session("set10_colour") = "1" then Response.Write " selected" end if%> value="1">Colour 1</option>
                              <option <% if session("set10_colour") = "2" then Response.Write " selected" end if%> value="2">Colour 2</option>
                              <option <% if session("set10_colour") = "3" then Response.Write " selected" end if%> value="3">Colour 3</option>
                              <option <% if session("set10_colour") = "4" then Response.Write " selected" end if%> value="4">Colour 4</option>
                              <option <% if session("set10_colour") = "5" then Response.Write " selected" end if%> value="5">Colour 5</option>
                            </select></td>
                          <td><input type="text" id="txtItem10" name="txtItem10" maxlength="15" size="20" value="<%= session("set10") %>" /></td>
                          <td><input type="text" id="txtQty10" name="txtQty10" maxlength="4" size="5" value="<%= session("set10_qty") %>" /></td>
                        </tr>
                      </table></td>
                  </tr>
                </table>
                <p id="kit_note"><strong>Kit:</strong> Created at TT, items to be placed into pre made cartons and all components placed into one outer carton.</p>
                <p id="set_note"><strong>Set:</strong> Multi box items which are picked individually and sent as individual components.</p>
                <table cellpadding="3" cellspacing="0" class="item_maintenance_box">
                  <tr>
                    <td class="item_maintenance_header" colspan="4">4. Dimensions</td>
                  </tr>
                  <tr>
                    <td width="25%">Gross weight<span class="mandatory">*</span>: </td>
                    <td width="25%"><input type="text" id="txtGrossWeight" name="txtGrossWeight" maxlength="5" size="5" value="<%= session("gross_weight") %>" />
                      kg</td>
                    <td width="20%">Nett weight<span class="mandatory">*</span>:</td>
                    <td width="30%"><input type="text" id="txtNettWeight" name="txtNettWeight" maxlength="5" size="5" value="<%= session("nett_weight") %>" />
                      kg</td>
                  </tr>
                  <tr>
                    <td>Width<span class="mandatory">*</span>:</td>
                    <td><input type="text" id="txtWidth" name="txtWidth" maxlength="5" size="5" value="<%= session("width") %>" />
                      cm</td>
                    <td>&nbsp;</td>
                    <td>&nbsp;</td>
                  </tr>
                  <tr>
                    <td>Height<span class="mandatory">*</span>:</td>
                    <td><input type="text" id="txtHeight" name="txtHeight" maxlength="5" size="5" value="<%= session("height") %>" />
                      cm</td>
                    <td>&nbsp;</td>
                    <td>&nbsp;</td>
                  </tr>
                  <tr>
                    <td>Depth<span class="mandatory">*</span>:</td>
                    <td><input type="text" id="txtDepth" name="txtDepth" maxlength="5" size="5" value="<%= session("depth") %>" />
                      cm</td>
                    <td>&nbsp;</td>
                    <td>&nbsp;</td>
                  </tr>
                  <tr>
                    <td>Volume:</td>
                    <td><%= left(session("volume"),5) %> m<sup>3</sup></td>
                    <td>&nbsp;</td>
                    <td>&nbsp;</td>
                  </tr>
                  <tr>
                    <td>Packaging qty:</td>
                    <td><input type="text" id="txtPackUnit" name="txtPackUnit" maxlength="5" size="5" value="<%= session("pack_unit") %>" /></td>
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
                    <td>Vendor:</td>
                    <td><input type="text" id="txtVendor" name="txtVendor" maxlength="20" size="20" value="<%= session("vendor") %>" /></td>
                  </tr>
                  <tr>
                    <td>Country of origin:</td>
                    <td><select name="cboCountryOrigin">
                        <%= strCountryList %>
                      </select></td>
                  </tr>
                  <tr>
                    <td>EAN code:</td>
                    <td><input type="text" id="txtEANcode" name="txtEANcode" maxlength="30" size="40" value="<%= session("ean_code") %>" /></td>
                  </tr>
                  <tr>
                    <td>Created in BASE:</td>
                    <td>
                      <input type="button" id="btnLogisticsPending" value="Done" <% if session("logistics_pending") = "True" then Response.Write "disabled" end if %>>
                      <div id="success" class="label green_font">Success</div>
                      <div id="error" class="label red_font">Error</div>
                    </td>
                  </tr>
                </table></td>
              <td width="33%" valign="top"><table cellpadding="3" cellspacing="0" class="item_maintenance_box" id="pricing">
                  <tr>
                    <td class="item_maintenance_header" colspan="2">6. Pricing</td>
                  </tr>
                  <tr>
                    <td>FOB (P01):</td>
                    <td>$
                      <input type="text" id="txtFOB" name="txtFOB" maxlength="10" size="10" value="<%= session("fob") %>" /></td>
                  </tr>
                  <tr>
                    <td width="30%">Trade (S01 ex tax)<span class="mandatory">*</span>:</td>
                    <td width="70%">$
                      <input type="text" id="txtTrade" name="txtTrade" maxlength="10" size="10" value="<%= session("trade") %>" /></td>
                  </tr>
                  <tr>
                    <td>RRP (S50 inc tax)<span class="mandatory">*</span>:</td>
                    <td>$
                      <input type="text" id="txtRRP" name="txtRRP" maxlength="10" size="10" value="<%= session("rrp") %>" /></td>
                  </tr>
                  <tr>
                    <td>NZ Trade (S02 ex tax):</td>
                    <td>$
                      <input type="text" id="txtNZtrade" name="txtNZtrade" maxlength="10" size="10" value="<%= session("nz_trade") %>" /></td>
                  </tr>
                  <tr>
                    <td>Mod required?</td>
                    <td><select name="cboModRequired">
                        <option <% if session("mod_required") = "0" then Response.Write " selected" end if%> value="0">No</option>
                        <option <% if session("mod_required") = "1" then Response.Write " selected" end if%> value="1">Yes</option>
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
                <table cellpadding="3" cellspacing="0" class="item_maintenance_box" id="additional_info">
                  <tr>
                    <td class="item_maintenance_header">7. Additional Info</td>
                  </tr>
                  <tr>
                    <td><textarea name="txtComments" id="txtComments" cols="50" rows="8" onKeyDown="limitText(this.form.txtComments,this.form.countdown,200);" onKeyUp="limitText(this.form.txtComments,this.form.countdown,200);"><%= session("comments") %></textarea></td>
                  </tr>
                  <tr class="status_row">
                    <td>Status:
                      <select name="cboStatus">
                        <option <% if session("status") = "1" then Response.Write " selected" end if%> value="1">Open</option>
                        <% if session("gm_approval") = "1" and session("emc_approval") = "1" and session("logistics_processed") = "1" then %>
                        <option <% if session("status") = "0" then Response.Write " selected" end if%> value="0">Completed</option>
                        <% end if %>
                    </select></td>
                  </tr>
                </table></td>
            </tr>
          </table>
        </form>
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
        </table>
      </div></td>
  </tr>
</table>

<script type="text/javascript">
$(function() {

    // Role 3 (Mondiale) should only see Dimensions and Logistics boxes
    if ($("#UsrLoginRole").val() === "3") {
        $("#item_details").hide();
        $("#general").hide();
        $("#set_item_kit_item").hide();
        $("#kit_note").hide();
        $("#set_note").hide();
        $("#pricing").hide();
        $("#additional_info").hide();
    }

    // Hide the 'Created in BASE' success and error displays
    $("#success").hide()
    $("#error").hide()

    $("#btnLogisticsPending").click(function(e) {
        // Cancel the postback
        e.preventDefault();

        // Get the item_id
        var data = getUrlVars()["id"];

        // Send AJAX request to server to flag the logistics_pending field for this item_id
        $.ajax({
            url: 'update_item-maintenance-base.asp',
            type: 'POST',
            data: 'item_id=' + data,
            success: function(successMsg) {
                if (successMsg === 'SUCCESS') {
                    // All went well
                    $("#success").fadeIn(400).delay(4000).fadeOut(400);

                    // Disable the button
                    $("#btnLogisticsPending").attr("disabled", true);
                } else {
                    // Something went wrong
                    $("#error").fadeIn(400).delay(4000).fadeOut(400);
                }
            },
            error: function(errorMsg) {
                // Something went wrong
                $("#error").fadeIn(400).delay(4000).fadeOut(400);
            }
        });
    });
});

// Read a page's GET URL variables and return them as an associative array.
function getUrlVars() {
    var vars = [], hash;
    var hashes = window.location.href.slice(window.location.href.indexOf('?') + 1).split('&');

    for(var i = 0; i < hashes.length; i++) {
        hash = hashes[i].split('=');
        vars.push(hash[0]);
        vars[hash[0]] = hash[1];
    }

    return vars;
}
</script>

</body>
</html>