<!--#include file="include/connection_it.asp " -->
<% strSection = "roadshow" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Add Roadshow Item</title>
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="include/generic_form_validations.js"></script>
<script language="JavaScript" type="text/javascript">
function validateFormOnSubmit(theForm) {
	var reason = "";
	var blnSubmit = true;
	
	reason += validateEmptyField(theForm.txtProductCode,"Product Code");
	//reason += validateSpecialCharacters(theForm.txtProductCode,"Product Code");
	
	reason += validateEmptyField(theForm.txtDescription,"Description");
	reason += validateSpecialCharacters(theForm.txtDescription,"Description");
	
	reason += validateNumeric(theForm.txtQuantity,"Quantity");
	
  	if (reason != "") {
    	alert("Some fields need correction:\n" + reason);
    	
		blnSubmit = false;
		return false;
  	}

	if (blnSubmit == true){
		//alert("Yea");
        theForm.Action.value = 'Add';
  		//theForm.submit();
		
		return true;
    }
}

</script>
<%
'Call OpenDataBase

sub addItem
	dim strSQL

	call OpenDataBase()
		
	strSQL = "INSERT INTO yma_roadshow_2014 ("
	strSQL = strSQL & " product_code, "
	strSQL = strSQL & " item_group, "
	strSQL = strSQL & " description, "
	strSQL = strSQL & " category, "
	strSQL = strSQL & " department, "
	strSQL = strSQL & " rrp, "
	strSQL = strSQL & " sku_type, "
	strSQL = strSQL & " prototype, "
	strSQL = strSQL & " quantity, "
	strSQL = strSQL & " packaging, "
	strSQL = strSQL & " source, "
	strSQL = strSQL & " origin, "
	strSQL = strSQL & " available, "
	strSQL = strSQL & " transit, "
	strSQL = strSQL & " invoice_no, "
	strSQL = strSQL & " pallet_no, "
	strSQL = strSQL & " loading_sequence, "
	strSQL = strSQL & " type, "
	strSQL = strSQL & " available_for_sale, "
	strSQL = strSQL & " pre_sold, "
	strSQL = strSQL & " comments, "
	strSQL = strSQL & " return_to, "
	strSQL = strSQL & " displayed, "
	strSQL = strSQL & " how_displayed, "
	strSQL = strSQL & " fb_completed, "	
	strSQL = strSQL & " date_created, "
	strSQL = strSQL & " created_by) VALUES ( "
	strSQL = strSQL & "'" & Server.HTMLEncode(Request.Form("txtProductCode")) & "',"	
	strSQL = strSQL & "'" & Trim(Request.Form("txtItemGroup")) & "',"	
	strSQL = strSQL & "'" & Server.HTMLEncode(Request.Form("txtDescription")) & "',"	
	strSQL = strSQL & "'" & Trim(Request.Form("cboCategory")) & "',"	
	strSQL = strSQL & "'" & Trim(Request.Form("cboDepartment")) & "',"
	strSQL = strSQL & "'" & Trim(Request.Form("txtRRP")) & "',"
	strSQL = strSQL & "'" & Trim(Request.Form("cboSkuType")) & "',"
	strSQL = strSQL & "'" & Trim(Request.Form("cboPrototype")) & "',"
	strSQL = strSQL & "'" & Trim(Request.Form("txtQuantity")) & "',"
	strSQL = strSQL & "'" & Trim(Request.Form("cboPackaging")) & "',"	
	strSQL = strSQL & "'" & Trim(Request.Form("cboSource")) & "',"		
	strSQL = strSQL & "'" & Trim(Request.Form("cboOrigin")) & "',"	
	strSQL = strSQL & "'" & Trim(Request.Form("cboAvailable")) & "',"
	strSQL = strSQL & "'" & Trim(Request.Form("cboTransit")) & "',"
	strSQL = strSQL & "'" & Trim(Request.Form("txtInvoiceNo")) & "',"
	strSQL = strSQL & "'" & Trim(Request.Form("txtPalletNo")) & "',"
	strSQL = strSQL & "'" & Trim(Request.Form("txtLoadingSequence")) & "',"
	strSQL = strSQL & "'" & Trim(Request.Form("cboType")) & "',"
	strSQL = strSQL & "'" & Trim(Request.Form("cboAvailableForSale")) & "',"
	strSQL = strSQL & "'" & Trim(Request.Form("cboPreSold")) & "',"
	strSQL = strSQL & "'" & Replace(Request.Form("txtComment"),"'","''") & "',"
	strSQL = strSQL & "'" & Trim(Request.Form("cboReturnTo")) & "',"	
	strSQL = strSQL & "'" & Trim(Request.Form("txtDisplay")) & "',"
	strSQL = strSQL & "'" & Trim(Request.Form("cboHowDisplayed")) & "',"
	strSQL = strSQL & "'" & Trim(Request.Form("cboFB")) & "',"	
	strSQL = strSQL & "getdate(),"
	strSQL = strSQL & "'" & session("UsrUserName") & "')"
	
	response.Write strSQL	  
	on error resume next
	conn.Execute strSQL
	
	'On error Goto 0  
	
	if err <> 0 then
		strMessageText = err.description
	else 		
		Response.Redirect("thank-you_roadshow.asp")
	end if 
	
	Call CloseDataBase()
end sub

sub main
	call UTL_validateLogin
	
	if Request.ServerVariables("REQUEST_METHOD") = "POST" then
		if Trim(Request("Action")) = "Add" then		
			call addItem
		end if
	end if
end sub

call main
%>
</head>
<body>
<table width="100%" cellpadding="0" cellspacing="0">
  <!-- #include file="include/header.asp" -->
  <tr>
    <td class="first_content"><img src="images/backward_arrow.gif" border="0" /> <a href="list_roadshow.asp">Back to List</a>
      <h2>Add NEW Roadshow Item</h2>
      <font color="green"><%= strMessageText %></font>
      <form action="" method="post" name="form_add_roadshow" id="form_add_roadshow" onsubmit="return validateFormOnSubmit(this)">
        <table border="0" cellpadding="5" cellspacing="0" width="1024">
          <tr>
            <td width="50%" valign="top"><table cellpadding="5" cellspacing="0" class="item_maintenance_box">
                <tr>
                  <td colspan="2" class="item_maintenance_header">Item Details</td>
                </tr>
                <tr>
                  <td width="25%" align="right">Owner:</td>
                  <td width="75%"><%= session("UsrUserName") %>
                  </td>
                </tr>
                <tr>
                  <td align="right">Dept:</td>
                  <td><select name="cboDepartment">
                      <option value="CA">CA</option>
                      <option value="Pro">Pro</option>
                      <option value="Trad">Trad</option>
                    </select></td>
                </tr>
                <tr>
                  <td align="right">Group:</td>
                  <td><input type="text" id="txtItemGroup" name="txtItemGroup" maxlength="5" size="5" /></td>
                </tr>
                <tr>
                  <td align="right">Category:</td>
                  <td><select name="cboCategory">
                      <option value="Acoustic Drums">Acoustic Drums</option>
                      <option value="Brass &amp; Woodwind">Brass &amp; Woodwind</option>
                      <option value="CA">CA</option>
                      <option value="Digital Pianos">Digital Pianos</option>
                      <option value="Electronic Drums">Electronic Drums</option>
                      <option value="Guitars">Guitars</option>
                      <option value="MPP">MPP</option>
                      <option value="Paiste">Paiste</option>
                      <option value="Percussion & Strings">Percussion & Strings</option>
                      <option value="Pianos">Pianos</option>                      
                      <option value="Portable Keyboards">Portable Keyboards</option>
                      <option value="Pro Audio">Pro Audio</option>
                      <option value="Syde">SYDE</option>
                      <option value="Vox">VOX</option>
                      
                  </select></td>
                </tr>
                <tr>
                  <td align="right">Product code<span class="mandatory">*</span>:</td>
                  <td><input type="text" id="txtProductCode" name="txtProductCode" maxlength="30" size="30" /></td>
                </tr>
                
                <tr>
                  <td align="right">Description<span class="mandatory">*</span>:</td>
                  <td><input type="text" id="txtDescription" name="txtDescription" maxlength="65" size="60" /></td>
                </tr>
                
                
                <tr>
                  <td align="right">RRP:</td>
                  <td>$
                  <input type="text" id="txtRRP" name="txtRRP" maxlength="8" size="8" /></td>
                </tr>
                <tr>
                  <td align="right">Qty<span class="mandatory">*</span>:</td>
                  <td><input type="text" id="txtQuantity" name="txtQuantity" maxlength="4" size="4" /></td>
                </tr>
                <tr>
                  <td align="right">SKU:</td>
                  <td><select name="cboSkuType">
                      <option value="A Sku">A Sku</option>
                      <option value="B Sku">B Sku</option>
                      <option value="New">New</option>
                    </select></td>
                </tr>
                <tr>
                  <td align="right">Prototype:</td>
                  <td><select name="cboPrototype">
                      <option value="No">No</option>
                      <option value="Yes">Yes</option>                      
                    </select></td>
                </tr>
                
                <tr>
                  <td align="right">Packaging:</td>
                  <td><select name="cboPackaging">
                      <option value="Carton">Carton</option>
                      <option value="Roadcase">Roadcase</option>
                      <option value="Unboxed">Unboxed</option>
                    </select></td>
                </tr>
                <tr>
                  <td align="right">Source:</td>
                  <td><select name="cboSource">
                      <option value="New Loan">New Loan</option>
                      <option value="Asset">Asset</option>
                      <option value="Loan Stock">Loan Stock</option>
                      <option value="New Pick">New Pick</option>
                      <option value="Promo Stock">Promo Stock</option>
                      <option value="3S">3S</option>
                  </select></td>
                </tr>
                
            </table></td>
            <td width="50%" valign="top"><table cellpadding="5" cellspacing="0" class="item_maintenance_box">
                <tr>
                  <td colspan="2" class="item_maintenance_header">Logistics Details</td>
                </tr>
                
                <tr>
                  <td width="30%" align="right">Origin:</td>
                  <td width="70%"><select name="cboOrigin">
                      <option value="3T">3T</option>
                      <option value="Direct from YCJ">Direct from YCJ</option>
                      <option value="Head Office">Head Office</option>
                    </select></td>
                </tr>
                <tr>
                  <td align="right">Available?</td>
                  <td><select name="cboAvailable">                      
                      <option value="Yes">Yes</option>
                      <option value="No">No</option>
                      <option value="In Transit">In Transit</option>
                    </select></td>
                </tr>
                <tr>
                  <td align="right">Transit:</td>
                  <td><select name="cboTransit">
                      <option value="NA">NA</option>
                      <option value="Air Freight">Air Freight</option>
                      <option value="Sea Freight">Sea Freight</option>                      
                    </select></td>
                </tr>
                
                <tr>
                  <td align="right">Type:</td>
                  <td><select name="cboType">
                      <option value="Existing Asset">Existing Asset</option>
                      <option value="Existing Loan">Existing Loan</option>
                      <option value="New Loan">New Loan</option>
                      <option value="New Order">New Order</option>
                    </select></td>
                </tr>
                <tr>
                  <td align="right">Available for sale?</td>
                  <td><select name="cboAvailableForSale">
                      <option value="Yes">Yes</option>
                      <option value="No">No</option>                      
                    </select></td>
                </tr>
                <tr>
                  <td align="right">Pre-sold?</td>
                  <td><select name="cboPreSold">
                      <option value="No">No</option>
                      <option value="Yes">Yes</option>
                    </select></td>
                </tr>
                <tr>
                  <td align="right">After-show Return to:</td>
                  <td><select name="cboReturnTo">
                      <option value="Head Office">Head Office</option>
                      <option value="3S">3S</option>
                      <option value="3T">3T</option>
                      <option value="Dealer">Dealer</option>                      
                    </select></td>
                </tr>
                <tr>
                  <td align="right">Display at:</td>
                  <td><input type="text" id="txtDisplay" name="txtDisplay" maxlength="30" size="30" /></td>
                </tr>
                <tr>
                  <td align="right">How displayed?</td>
                  <td><select name="cboHowDisplayed">
                      <option value="Floor">Floor</option>
                      <option value="Slatwall">Slatwall</option>
                      <option value="Table">Table</option>
                    </select></td>
                </tr>
                <tr>
                  <td align="right">F &amp; B completed?</td>
                  <td><select name="cboFB">
                      <option value="Yes">Yes</option>
                      <option value="No">No</option>                      
                    </select></td>
                </tr>
                <tr>
                  <td align="right">Comments:</td>
                  <td><textarea name="txtComment" id="txtComment" cols="35" rows="4"></textarea></td>
                </tr>     
                <tr class="status_row">
                  <td align="right">Invoice no:</td>
                  <td><input type="text" id="txtInvoiceNo" name="txtInvoiceNo" maxlength="30" size="30" /></td>
                </tr>
                <tr class="status_row">
                  <td align="right">Pallet no:</td>
                  <td><input type="text" id="txtPalletNo" name="txtPalletNo" maxlength="30" size="30" /></td>
                </tr>
                <tr class="status_row">
                  <td align="right">Loading sequence:</td>
                  <td><input type="text" id="txtLoadingSequence" name="txtLoadingSequence" maxlength="30" size="30" /></td>
                </tr>           
              </table>
              <p>
                <input type="hidden" name="Action" />
                <input type="submit" value="Add Roadshow Item" />
              </p></td>
          </tr>
        </table>
      </form></td>
  </tr>
  
</table>
</body>
</html>