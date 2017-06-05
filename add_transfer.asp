<!--#include file="include/connection_it.asp " -->
<% strSection = "transfer" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Add Transfer</title>
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="include/generic_form_validations.js"></script>
<script language="JavaScript" type="text/javascript">   
function validateFormOnSubmit(theForm) {
	var reason = "";
	var blnSubmit = true;

	reason += validateEmptyField(theForm.txtProduct1,"Product 1");
	reason += validateSpecialCharacters(theForm.txtProduct1,"Product 1");

	reason += validateEmptyField(theForm.txtQty1,"Qty 1");
	reason += validateSpecialCharacters(theForm.txtQty1,"Qty 1");

	reason += validateEmptyField(theForm.txtPallet1,"Pallet 1");
	reason += validateSpecialCharacters(theForm.txtPallet1,"Pallet 1");

	reason += validateEmptyField(theForm.txtDeliveryDate,"Delivery Date");

	reason += validateEmptyField(theForm.txtDeliveryTime,"Delivery Time");
	reason += validateSpecialCharacters(theForm.txtDeliveryTime,"Delivery Time");

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
sub addTransfer
	dim strSQL

	dim strCreatedBy
	dim strPriority

	dim strWarehouse
	dim strNoPallets
	dim strProduct1
	dim intQty1
	dim intPallet1
	dim strInfo1
	dim strProduct2
	dim intQty2
	dim intPallet2
	dim strInfo2
	dim strProduct3
	dim intQty3
	dim intPallet3
	dim strInfo3
	dim strProduct4
	dim intQty4
	dim intPallet4
	dim strInfo4
	dim strProduct5
	dim intQty5
	dim intPallet5
	dim strInfo5
	dim strProduct6
	dim intQty6
	dim intPallet6
	dim strInfo6
	dim strProduct7
	dim intQty7
	dim intPallet7
	dim strInfo7
	dim strProduct8
	dim intQty8
	dim intPallet8
	dim strInfo8
	dim strProduct9
	dim intQty9
	dim intPallet9
	dim strInfo9
	dim strProduct10
	dim intQty10
	dim intPallet10
	dim strInfo10
	dim strProduct11
	dim intQty11
	dim intPallet11
	dim strInfo11
	dim strProduct12
	dim intQty12
	dim intPallet12
	dim strInfo12
	dim strProduct13
	dim intQty13
	dim intPallet13
	dim strInfo13
	dim strProduct14
	dim intQty14
	dim intPallet14
	dim strInfo14
	dim strProduct15
	dim intQty15
	dim intPallet15
	dim strInfo15
	dim strProduct16
	dim intQty16
	dim intPallet16
	dim strInfo16
	dim strProduct17
	dim intQty17
	dim intPallet17
	dim strInfo17
	dim strProduct18
	dim intQty18
	dim intPallet18
	dim strInfo18
	dim strProduct19
	dim intQty19
	dim intPallet19
	dim strInfo19
	dim strProduct20
	dim intQty20
	dim intPallet20
	dim strInfo20
	dim strDeliveryDate
	dim strDeliveryTime
	dim strInvoiceNo
	dim strComments

	strCreatedBy		= Trim(Request("txtCreatedBy"))
	strCreatedBy 		= replace(strCreatedBy,"'","''")

	strPriority 		= Trim(Request("cboPriority"))

	strWarehouse		= Trim(Request("cboWarehouse"))

	strNoPallets  		= Trim(Request("txtNoPallets"))
	strNoPallets 		= replace(strNoPallets,"'","''")

	strProduct1  		= Trim(Request("txtProduct1"))
	strProduct1 		= replace(strProduct1,"'","''")
	intQty1  		 	= Trim(Request("txtQty1"))
	intQty1 			= replace(intQty1,"'","''")
	intPallet1  		= Trim(Request("txtPallet1"))
	intPallet1 			= replace(intPallet1,"'","''")
	strInfo1	  		= replace(Trim(Request("txtInfo1")),"'","''")

	strProduct2    		= Trim(Request("txtProduct2"))
	strProduct2 		= replace(strProduct2,"'","''")
	intQty2  		 	= Trim(Request("txtQty2"))
	intQty2 			= replace(intQty2,"'","''")
	intPallet2  		= Trim(Request("txtPallet2"))
	intPallet2 			= replace(intPallet2,"'","''")
	strInfo2	  		= replace(Trim(Request("txtInfo2")),"'","''")

	strProduct3  		= Trim(Request("txtProduct3"))
	strProduct3 		= replace(strProduct3,"'","''")
	intQty3  		 	= Trim(Request("txtQty3"))
	intQty3 			= replace(intQty3,"'","''")
	intPallet3  		= Trim(Request("txtPallet3"))
	intPallet3 			= replace(intPallet3,"'","''")
	strInfo3	  		= replace(Trim(Request("txtInfo3")),"'","''")

	strProduct4			= Trim(Request("txtProduct4"))
	strProduct4 		= replace(strProduct4,"'","''")
	intQty4  		 	= Trim(Request("txtQty4"))
	intQty4 			= replace(intQty4,"'","''")
	intPallet4  		= Trim(Request("txtPallet4"))
	intPallet4 			= replace(intPallet4,"'","''")
	strInfo4	  		= replace(Trim(Request("txtInfo4")),"'","''")

	strProduct5			= Trim(Request("txtProduct5"))
	strProduct5 		= replace(strProduct5,"'","''")
	intQty5  		 	= Trim(Request("txtQty5"))
	intQty5 			= replace(intQty5,"'","''")
	intPallet5  		= Trim(Request("txtPallet5"))
	intPallet5 			= replace(intPallet5,"'","''")
	strInfo5	  		= replace(Trim(Request("txtInfo5")),"'","''")

	strProduct6			= Trim(Request("txtProduct6"))
	strProduct6 		= replace(strProduct6,"'","''")
	intQty6  		 	= Trim(Request("txtQty6"))
	intQty6 			= replace(intQty6,"'","''")
	intPallet6  		= Trim(Request("txtPallet6"))
	intPallet6 			= replace(intPallet6,"'","''")
	strInfo6	  		= replace(Trim(Request("txtInfo6")),"'","''")

	strProduct7			= Trim(Request("txtProduct7"))
	strProduct7 		= replace(strProduct7,"'","''")
	intQty7  		 	= Trim(Request("txtQty7"))
	intQty7 			= replace(intQty7,"'","''")
	intPallet7  		= Trim(Request("txtPallet7"))
	intPallet7 			= replace(intPallet7,"'","''")
	strInfo7  			= replace(Trim(Request("txtInfo7")),"'","''")

	strProduct8			= Trim(Request("txtProduct8"))
	strProduct8 		= replace(strProduct8,"'","''")
	intQty8  		 	= Trim(Request("txtQty8"))
	intQty8 			= replace(intQty8,"'","''")
	intPallet8  		= Trim(Request("txtPallet8"))
	intPallet8 			= replace(intPallet8,"'","''")
	strInfo8	  		= replace(Trim(Request("txtInfo8")),"'","''")

	strProduct9			= Trim(Request("txtProduct9"))
	strProduct9 		= replace(strProduct9,"'","''")
	intQty9  		 	= Trim(Request("txtQty9"))
	intQty9 			= replace(intQty9,"'","''")
	intPallet9  		= Trim(Request("txtPallet9"))
	intPallet9 			= replace(intPallet9,"'","''")
	strInfo9	  		= replace(Trim(Request("txtInfo9")),"'","''")

	strProduct10		= Trim(Request("txtProduct10"))
	strProduct10 		= replace(strProduct10,"'","''")
	intQty10  		 	= Trim(Request("txtQty10"))
	intQty10 			= replace(intQty10,"'","''")
	intPallet10  		= Trim(Request("txtPallet10"))
	intPallet10 		= replace(intPallet10,"'","''")
	strInfo10	  		= replace(Trim(Request("txtInfo10")),"'","''")

	strProduct11    	= Trim(Request("txtProduct11"))
	strProduct11 		= replace(strProduct11,"'","''")
	intQty11  		 	= Trim(Request("txtQty11"))
	intQty11 			= replace(intQty11,"'","''")
	intPallet11  		= Trim(Request("txtPallet11"))
	intPallet11 		= replace(intPallet11,"'","''")
	strInfo11	  		= replace(Trim(Request("txtInfo11")),"'","''")

	strProduct12    	= Trim(Request("txtProduct12"))
	strProduct12 		= replace(strProduct12,"'","''")
	intQty12  		 	= Trim(Request("txtQty12"))
	intQty12 			= replace(intQty12,"'","''")
	intPallet12  		= Trim(Request("txtPallet12"))
	intPallet12 		= replace(intPallet12,"'","''")
	strInfo12	  		= replace(Trim(Request("txtInfo12")),"'","''")

	strProduct13  		= Trim(Request("txtProduct13"))
	strProduct13 		= replace(strProduct13,"'","''")
	intQty13  		 	= Trim(Request("txtQty13"))
	intQty13 			= replace(intQty13,"'","''")
	intPallet13  		= Trim(Request("txtPallet13"))
	intPallet13 		= replace(intPallet13,"'","''")
	strInfo13  			= replace(Trim(Request("txtInfo13")),"'","''")

	strProduct14		= Trim(Request("txtProduct14"))
	strProduct14 		= replace(strProduct14,"'","''")
	intQty14  		 	= Trim(Request("txtQty14"))
	intQty14 			= replace(intQty14,"'","''")
	intPallet14  		= Trim(Request("txtPallet14"))
	intPallet14 		= replace(intPallet14,"'","''")
	strInfo14	  		= replace(Trim(Request("txtInfo14")),"'","''")

	strProduct15		= Trim(Request("txtProduct15"))
	strProduct15 		= replace(strProduct15,"'","''")
	intQty15  		 	= Trim(Request("txtQty15"))
	intQty15 			= replace(intQty15,"'","''")
	intPallet15  		= Trim(Request("txtPallet15"))
	intPallet15 		= replace(intPallet15,"'","''")
	strInfo15	  		= replace(Trim(Request("txtInfo15")),"'","''")

	strProduct16		= Trim(Request("txtProduct16"))
	strProduct16 		= replace(strProduct16,"'","''")
	intQty16  		 	= Trim(Request("txtQty16"))
	intQty16 			= replace(intQty16,"'","''")
	intPallet16  		= Trim(Request("txtPallet16"))
	intPallet16 		= replace(intPallet16,"'","''")
	strInfo16	  		= replace(Trim(Request("txtInfo16")),"'","''")

	strProduct17		= Trim(Request("txtProduct17"))
	strProduct17 		= replace(strProduct17,"'","''")
	intQty17  		 	= Trim(Request("txtQty17"))
	intQty17 			= replace(intQty17,"'","''")
	intPallet17  		= Trim(Request("txtPallet17"))
	intPallet17 		= replace(intPallet17,"'","''")
	strInfo17	  		= replace(Trim(Request("txtInfo17")),"'","''")

	strProduct18		= Trim(Request("txtProduct18"))
	strProduct18 		= replace(strProduct18,"'","''")
	intQty18  		 	= Trim(Request("txtQty18"))
	intQty18 			= replace(intQty18,"'","''")
	intPallet18  		= Trim(Request("txtPallet18"))
	intPallet18 		= replace(intPallet18,"'","''")
	strInfo18	  		= replace(Trim(Request("txtInfo18")),"'","''")

	strProduct19		= Trim(Request("txtProduct19"))
	strProduct19 		= replace(strProduct19,"'","''")
	intQty19  		 	= Trim(Request("txtQty19"))
	intQty19 			= replace(intQty19,"'","''")
	intPallet19  		= Trim(Request("txtPallet19"))
	intPallet19 		= replace(intPallet19,"'","''")
	strInfo19	  		= replace(Trim(Request("txtInfo19")),"'","''")

	strProduct20		= Trim(Request("txtProduct20"))
	strProduct20 		= replace(strProduct20,"'","''")
	intQty20  		 	= Trim(Request("txtQty20"))
	intQty20 			= replace(intQty20,"'","''")
	intPallet20  		= Trim(Request("txtPallet20"))
	intPallet20 		= replace(intPallet20,"'","''")
	strInfo20	  		= replace(Trim(Request("txtInfo20")),"'","''")

	strDeliveryDate		= Trim(Request("txtDeliveryDate"))
	strDeliveryTime		= Trim(Request("txtDeliveryTime"))
	strInvoiceNo		= Trim(Request("txtInvoiceNo"))

	strComments			= Trim(Request("txtComments"))
	strComments 		= replace(strComments,"'","''")

	Call OpenDataBase()

	strSQL = "INSERT INTO yma_transfer (priority, warehouse, product_1, qty_1, pallet_1, info_1 , product_2, qty_2, pallet_2, info_2 , product_3, qty_3, pallet_3, info_3 , product_4, qty_4, pallet_4, info_4 , product_5, qty_5, pallet_5, info_5 , product_6, qty_6, pallet_6, info_6 , product_7, qty_7, pallet_7, info_7 , product_8, qty_8, pallet_8, info_8 , product_9, qty_9, pallet_9, info_9 , product_10, qty_10, pallet_10, info_10 , product_11, qty_11, pallet_11, info_11 , product_12, qty_12, pallet_12, info_12 , product_13, qty_13, pallet_13, info_13 , product_14, qty_14, pallet_14, info_14 , product_15, qty_15, pallet_15, info_15 , product_16, qty_16, pallet_16, info_16 , product_17, qty_17, pallet_17, info_17 , product_18, qty_18, pallet_18, info_18 , product_19, qty_19, pallet_19, info_19 , product_20, qty_20, pallet_20, info_20, pickup_date, pickup_time, invoice_no, transfer_comments, created_by, status, date_created) VALUES ("
	strSQL = strSQL & "'" & strPriority & "',"
	strSQL = strSQL & "'" & strWarehouse & "',"
	strSQL = strSQL & "'" & strProduct1 & "',"
	strSQL = strSQL & "'" & intQty1 & "',"
	strSQL = strSQL & "'" & intPallet1 & "',"
	strSQL = strSQL & "'" & strInfo1 & "',"
	strSQL = strSQL & "'" & strProduct2 & "',"
	strSQL = strSQL & "'" & intQty2 & "',"
	strSQL = strSQL & "'" & intPallet2 & "',"
	strSQL = strSQL & "'" & strInfo2 & "',"
	strSQL = strSQL & "'" & strProduct3 & "',"
	strSQL = strSQL & "'" & intQty3 & "',"
	strSQL = strSQL & "'" & intPallet3 & "',"
	strSQL = strSQL & "'" & strInfo3 & "',"
	strSQL = strSQL & "'" & strProduct4 & "',"
	strSQL = strSQL & "'" & intQty4 & "',"
	strSQL = strSQL & "'" & intPallet4 & "',"
	strSQL = strSQL & "'" & strInfo4 & "',"
	strSQL = strSQL & "'" & strProduct5 & "',"
	strSQL = strSQL & "'" & intQty5 & "',"
	strSQL = strSQL & "'" & intPallet5 & "',"
	strSQL = strSQL & "'" & strInfo5 & "',"
	strSQL = strSQL & "'" & strProduct6 & "',"
	strSQL = strSQL & "'" & intQty6 & "',"
	strSQL = strSQL & "'" & intPallet6 & "',"
	strSQL = strSQL & "'" & strInfo6 & "',"
	strSQL = strSQL & "'" & strProduct7 & "',"
	strSQL = strSQL & "'" & intQty7 & "',"
	strSQL = strSQL & "'" & intPallet7 & "',"
	strSQL = strSQL & "'" & strInfo7 & "',"
	strSQL = strSQL & "'" & strProduct8 & "',"
	strSQL = strSQL & "'" & intQty8 & "',"
	strSQL = strSQL & "'" & intPallet8 & "',"
	strSQL = strSQL & "'" & strInfo8 & "',"
	strSQL = strSQL & "'" & strProduct9 & "',"
	strSQL = strSQL & "'" & intQty9 & "',"
	strSQL = strSQL & "'" & intPallet9 & "',"
	strSQL = strSQL & "'" & strInfo9 & "',"
	strSQL = strSQL & "'" & strProduct10 & "',"
	strSQL = strSQL & "'" & intQty10 & "',"
	strSQL = strSQL & "'" & intPallet10 & "',"
	strSQL = strSQL & "'" & strInfo10 & "',"
	strSQL = strSQL & "'" & strProduct11 & "',"
	strSQL = strSQL & "'" & intQty11 & "',"
	strSQL = strSQL & "'" & intPallet11 & "',"
	strSQL = strSQL & "'" & strInfo11 & "',"
	strSQL = strSQL & "'" & strProduct12 & "',"
	strSQL = strSQL & "'" & intQty12 & "',"
	strSQL = strSQL & "'" & intPallet12 & "',"
	strSQL = strSQL & "'" & strInfo12 & "',"
	strSQL = strSQL & "'" & strProduct13 & "',"
	strSQL = strSQL & "'" & intQty13 & "',"
	strSQL = strSQL & "'" & intPallet13 & "',"
	strSQL = strSQL & "'" & strInfo13 & "',"
	strSQL = strSQL & "'" & strProduct14 & "',"
	strSQL = strSQL & "'" & intQty14 & "',"
	strSQL = strSQL & "'" & intPallet14 & "',"
	strSQL = strSQL & "'" & strInfo14 & "',"
	strSQL = strSQL & "'" & strProduct15 & "',"
	strSQL = strSQL & "'" & intQty15 & "',"
	strSQL = strSQL & "'" & intPallet15 & "',"
	strSQL = strSQL & "'" & strInfo15 & "',"
	strSQL = strSQL & "'" & strProduct16 & "',"
	strSQL = strSQL & "'" & intQty16 & "',"
	strSQL = strSQL & "'" & intPallet16 & "',"
	strSQL = strSQL & "'" & strInfo16 & "',"
	strSQL = strSQL & "'" & strProduct17 & "',"
	strSQL = strSQL & "'" & intQty17 & "',"
	strSQL = strSQL & "'" & intPallet17 & "',"
	strSQL = strSQL & "'" & strInfo17 & "',"
	strSQL = strSQL & "'" & strProduct18 & "',"
	strSQL = strSQL & "'" & intQty18 & "',"
	strSQL = strSQL & "'" & intPallet18 & "',"
	strSQL = strSQL & "'" & strInfo18 & "',"
	strSQL = strSQL & "'" & strProduct19 & "',"
	strSQL = strSQL & "'" & intQty19 & "',"
	strSQL = strSQL & "'" & intPallet19 & "',"
	strSQL = strSQL & "'" & strInfo19 & "',"
	strSQL = strSQL & "'" & strProduct20 & "',"
	strSQL = strSQL & "'" & intQty20 & "',"
	strSQL = strSQL & "'" & intPallet20 & "',"
	strSQL = strSQL & "'" & strInfo20 & "',"
	strSQL = strSQL & " CONVERT(datetime,'" & strDeliveryDate & "',103),"
	strSQL = strSQL & "'" & strDeliveryTime & "',"
	strSQL = strSQL & "'" & strInvoiceNo & "',"
	strSQL = strSQL & "'" & strComments & "',"
	strSQL = strSQL & "'" & session("UsrUserName") & "',1, getdate())"

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

	emailFrom 		= "automailer@music.yamaha.com"

	on error resume next
	conn.Execute strSQL

	'On error Goto 0

	if err <> 0 then
		strMessageText = err.description
	else
		Select Case strWarehouse
			case "3K - TT"
				emailTo = "YMA-Warehouse@ttlogistics.com.au"
			case "TT - 3K"
				emailTo = "nicole.aquilina@silklogistics.com.au"
			case "3K - Excel"
				emailTo = "Yamaha_returns@exceltechnology.com.au"
				emailCc = "tonyk@exceltechnology.com.au"
			case "Excel - 3K"
				emailTo = "nicole.aquilina@silklogistics.com.au"
			case "TT - Excel"
				emailTo = "Yamaha_returns@exceltechnology.com.au"
				emailCc = "tonyk@exceltechnology.com.au"
			case "Excel - TT"
				emailTo = "YMA-Warehouse@ttlogistics.com.au"
			case "YMA - TT"
				emailTo = "YMA-Warehouse@ttlogistics.com.au"
			case "3L - TT"
				emailTo = "YMA-Warehouse@ttlogistics.com.au"
			case "TT - YMA"
				emailTo = "logistics-aus@music.yamaha.com"
			case "Excel - 3H"
				emailTo = "logistics-aus@music.yamaha.com"
				emailCc = "tonyk@exceltechnology.com.au"
			case "3XL - 3HX"
				emailTo = "logistics-aus@music.yamaha.com"
				emailCc = "tonyk@exceltechnology.com.au"
			case "3T - 3HT"
				emailTo = "logistics-aus@music.yamaha.com"
				emailCc = "YMA-Warehouse@ttlogistics.com.au"
            case "3H - 3HT", "3HT - 3H"
                emailTo = "logistics-aus@music.yamaha.com"
            case "REP - 3XL"
                emailTo = "tonyk@exceltechnology.com.au"
            case "TT - YME"
                emailTo = "carolyn.simonds@music.yamaha.com"
                emailCc = "YMA-Warehouse@ttlogistics.com.au"
			case "3L - 3XL"
                emailTo = "tonyk@exceltechnology.com.au"
				emailTo = "logistics-aus@music.yamaha.com"
			case "3L - 3S"
                emailTo = "tonyk@exceltechnology.com.au"
				emailTo = "logistics-aus@music.yamaha.com"
		end select

		emailBcc = "logistics-aus@music.yamaha.com"
		emailSubject = "(" & strWarehouse & ") Transfer Request: " & strDeliveryDate & "-" & strDeliveryTime &  " by: " & session("UsrUserName")

		emailBodyText =	"Requested by: " & session("UsrUserName") & vbCrLf _
					&	"---------------------------------------------------------------------------" & vbCrLf _
					&	"TRANSFER DETAILS" & vbCrLf _
					&	"---------------------------------------------------------------------------" & vbCrLf _
					&	"              Name / Qty / No of Pallet" & vbCrLf _
					&	"Product 1:    " & strProduct1 & "   " & intQty1 & "   " & intPallet1  & "   " & strInfo1 & vbCrLf _
					&	"Product 2:    " & strProduct2 & "   " & intQty2 & "   " & intPallet2  & "   " & strInfo2 & vbCrLf _
					&	"Product 3:    " & strProduct3 & "   " & intQty3 & "   " & intPallet3  & "   " & strInfo3 & vbCrLf _
					&	"Product 4:    " & strProduct4 & "   " & intQty4 & "   " & intPallet4  & "   " & strInfo4 & vbCrLf _
					&	"Product 5:    " & strProduct5 & "   " & intQty5 & "   " & intPallet5  & "   " & strInfo5 & vbCrLf _
					&	"Product 6:    " & strProduct6 & "   " & intQty6 & "   " & intPallet6  & "   " & strInfo6 & vbCrLf _
					&	"Product 7:    " & strProduct7 & "   " & intQty7 & "   " & intPallet7  & "   " & strInfo7 & vbCrLf _
					&	"Product 8:    " & strProduct8 & "   " & intQty8 & "   " & intPallet8  & "   " & strInfo8 & vbCrLf _
					&	"Product 9:    " & strProduct9 & "   " & intQty9 & "   " & intPallet9  & "   " & strInfo9 & vbCrLf _
					&	"Product 10:   " & strProduct10 & "  " & intQty10 & "  " & intPallet10  & "   " & strInfo10 & vbCrLf _
					&	"Product 11:   " & strProduct11 & "  " & intQty11 & "  " & intPallet11  & "   " & strInfo11 & vbCrLf _
					&	"Product 12:   " & strProduct12 & "  " & intQty12 & "  " & intPallet12  & "   " & strInfo12 & vbCrLf _
					&	"Product 13:   " & strProduct13 & "  " & intQty13 & "  " & intPallet13  & "   " & strInfo13 & vbCrLf _
					&	"Product 14:   " & strProduct14 & "  " & intQty14 & "  " & intPallet14  & "   " & strInfo14 & vbCrLf _
					&	"Product 15:   " & strProduct15 & "  " & intQty15 & "  " & intPallet15  & "   " & strInfo15 & vbCrLf _
					&	"Product 16:   " & strProduct16 & "  " & intQty16 & "  " & intPallet16  & "   " & strInfo16 & vbCrLf _
					&	"Product 17:   " & strProduct17 & "  " & intQty17 & "  " & intPallet17  & "   " & strInfo17 & vbCrLf _
					&	"Product 18:   " & strProduct18 & "  " & intQty18 & "  " & intPallet18  & "   " & strInfo18 & vbCrLf _
					&	"Product 19:   " & strProduct19 & "  " & intQty19 & "  " & intPallet19  & "   " & strInfo19 & vbCrLf _
					&	"Product 20:   " & strProduct20 & "  " & intQty20 & "  " & intPallet20  & "   " & strInfo20 & vbCrLf _
					&	"Delivery Date :     " & strDeliveryDate & vbCrLf _
					&	"Delivery Time :     " & strDeliveryTime & vbCrLf _
					&	"Invoice No    :     " & strInvoiceNo & vbCrLf _
					&	"Comments      :     " & strComments & vbCrLf _
					&	"---------------------------------------------------------------------------" & vbCrLf _
					&	" " & vbCrLf _
					&   "This is an automated email - please do not reply to this email."

		Set oMail.Configuration = iConf
		oMail.To 		= emailTo
		oMail.Cc		= emailCc
		oMail.Bcc		= emailBcc
		oMail.From 		= emailFrom
		oMail.Subject 	= emailSubject
		oMail.TextBody 	= emailBodyText
		oMail.Send

		Set iConf = Nothing
		Set Flds = Nothing

		'Response.Write "Email sent"
		Response.Redirect("thank-you_transfer.asp")
	end if

	Call CloseDataBase()

end sub

sub main
	if Trim(Request("Action")) = "Add" then
		call addTransfer
	end if
end sub

dim strDisplayList

call main

dim strMessageText
%>
</head>
<body>
<form action="" method="post" name="form_add_transfer" id="form_add_transfer" onsubmit="return validateFormOnSubmit(this)">
  <table width="100%" cellpadding="0" cellspacing="0">
    <!-- #include file="include/header.asp" -->
    <tr>
      <td class="first_content"><table cellpadding="5" cellspacing="0" border="0">
          <tr>
            <td><a href="list_transfer.asp"><img src="images/icon_transfer.jpg" border="0" alt="Transfer Requests" /></a></td>
            <td valign="top"><img src="images/backward_arrow.gif" border="0" /> <a href="list_transfer.asp">Back to List</a>
              <h2>Add New Transfer</h2>
              <font color="green"><%= strMessageText %></font></td>
          </tr>
        </table>
       <table cellpadding="5" cellspacing="0" class="item_maintenance_box">
          <tr>
            <td colspan="2" class="item_maintenance_header">Transfer Details</td>
          </tr>
          <tr>
            <td width="20%">From - To:</td>
            <td width="80%"><select name="cboWarehouse">
                <option value="3K - TT">3K - TT</option>
                <option value="TT - 3K">TT - 3K</option>
                <option value="3K - Excel">3K - Excel</option>
                <option value="Excel - 3K">Excel - 3K</option>
                <option value="TT - Excel">TT - Excel</option>
                <option value="Excel - TT">Excel - TT</option>
                <option value="YMA - TT">YMA - TT</option>
                <option value="3L - TT">3L - TT</option>
                <option value="TT - YMA">TT - YMA</option>
                <option value="3XL - 3H">3XL - 3H</option>
                <option value="3XL - 3ND">3XL - 3ND</option>
                <option value="3T - 3HT">3T - 3HT</option>
                <option value="3XL - 3HX">3XL - 3HX</option>
                <option value="3H - 3HT">3H - 3HT</option>
                <option value="3HT - 3H">3HT - 3H</option>
                <option value="REP - 3XL">REP - 3XL</option>
                <option valie="TT - YME">TT - YME</option>
				<option valie="3L - 3XL">3L - 3XL</option>
				<option valie="3L - 3S">3L - 3S</option>
              </select>
            </td>
          </tr>
          <tr>
            <td>Priority? <img src="../logistics/images/icon_priority.gif" border="0" /></td>
            <td><select name="cboPriority">
                <option value="0">No</option>
                <option value="1">Yes</option>
              </select></td>
          </tr>
          <tr>
            <td colspan="2"><table width="100%">
                <tr>
                  <td><table width="400">
                      <tr>
                        <td>&nbsp;</td>
                        <td>Product</td>
                        <td>Qty</td>
                        <td>Pallet(s)</td>
                        <td>Info</td>
                      </tr>
                      <tr>
                        <td>1<span class="mandatory">*</span>:</td>
                        <td><input type="text" id="txtProduct1" name="txtProduct1" maxlength="15" size="15" /></td>
                        <td><input type="text" id="txtQty1" name="txtQty1" maxlength="4" size="5" /></td>
                        <td><input type="text" id="txtPallet1" name="txtPallet1" maxlength="4" size="5" /></td>
                        <td><input type="text" id="txtInfo1" name="txtInfo1" maxlength="15" size="15" /></td>
                      </tr>
                      <tr>
                        <td>2:</td>
                        <td><input type="text" id="txtProduct2" name="txtProduct2" maxlength="15" size="15" /></td>
                        <td><input type="text" id="txtQty2" name="txtQty2" maxlength="4" size="5" /></td>
                        <td><input type="text" id="txtPallet2" name="txtPallet2" maxlength="4" size="5" /></td>
                        <td><input type="text" id="txtInfo2" name="txtInfo2" maxlength="15" size="15" /></td>
                      </tr>
                      <tr>
                        <td>3:</td>
                        <td><input type="text" id="txtProduct3" name="txtProduct3" maxlength="15" size="15" /></td>
                        <td><input type="text" id="txtQty3" name="txtQty3" maxlength="4" size="5" /></td>
                        <td><input type="text" id="txtPallet3" name="txtPallet3" maxlength="4" size="5" /></td>
                        <td><input type="text" id="txtInfo3" name="txtInfo3" maxlength="15" size="15" /></td>
                      </tr>
                      <tr>
                        <td>4:</td>
                        <td><input type="text" id="txtProduct4" name="txtProduct4" maxlength="15" size="15" /></td>
                        <td><input type="text" id="txtQty4" name="txtQty4" maxlength="4" size="5" /></td>
                        <td><input type="text" id="txtPallet4" name="txtPallet4" maxlength="4" size="5" /></td>
                        <td><input type="text" id="txtInfo4" name="txtInfo4" maxlength="15" size="15" /></td>
                      </tr>
                      <tr>
                        <td>5:</td>
                        <td><input type="text" id="txtProduct5" name="txtProduct5" maxlength="15" size="15" /></td>
                        <td><input type="text" id="txtQty5" name="txtQty5" maxlength="4" size="5" /></td>
                        <td><input type="text" id="txtPallet5" name="txtPallet5" maxlength="4" size="5" /></td>
                        <td><input type="text" id="txtInfo5" name="txtInfo5" maxlength="15" size="15" /></td>
                      </tr>
                      <tr>
                        <td>6:</td>
                        <td><input type="text" id="txtProduct6" name="txtProduct6" maxlength="15" size="15" /></td>
                        <td><input type="text" id="txtQty6" name="txtQty6" maxlength="4" size="5" /></td>
                        <td><input type="text" id="txtPallet6" name="txtPallet6" maxlength="4" size="5" /></td>
                        <td><input type="text" id="txtInfo6" name="txtInfo6" maxlength="15" size="15" /></td>
                      </tr>
                      <tr>
                        <td>7:</td>
                        <td><input type="text" id="txtProduct7" name="txtProduct7" maxlength="15" size="15" /></td>
                        <td><input type="text" id="txtQty7" name="txtQty7" maxlength="4" size="5" /></td>
                        <td><input type="text" id="txtPallet7" name="txtPallet7" maxlength="4" size="5" /></td>
                        <td><input type="text" id="txtInfo7" name="txtInfo7" maxlength="15" size="15" /></td>
                      </tr>
                      <tr>
                        <td>8:</td>
                        <td><input type="text" id="txtProduct8" name="txtProduct8" maxlength="15" size="15" /></td>
                        <td><input type="text" id="txtQty8" name="txtQty8" maxlength="4" size="5" /></td>
                        <td><input type="text" id="txtPallet8" name="txtPallet8" maxlength="4" size="5" /></td>
                        <td><input type="text" id="txtInfo8" name="txtInfo8" maxlength="15" size="15" /></td>
                      </tr>
                      <tr>
                        <td>9:</td>
                        <td><input type="text" id="txtProduct9" name="txtProduct9" maxlength="15" size="15" /></td>
                        <td><input type="text" id="txtQty9" name="txtQty9" maxlength="4" size="5" /></td>
                        <td><input type="text" id="txtPallet9" name="txtPallet9" maxlength="4" size="5" /></td>
                        <td><input type="text" id="txtInfo9" name="txtInfo9" maxlength="15" size="15" /></td>
                      </tr>
                      <tr>
                        <td>10:</td>
                        <td><input type="text" id="txtProduct10" name="txtProduct10" maxlength="15" size="15" /></td>
                        <td><input type="text" id="txtQty10" name="txtQty10" maxlength="4" size="5" /></td>
                        <td><input type="text" id="txtPallet10" name="txtPallet10" maxlength="4" size="5" /></td>
                        <td><input type="text" id="txtInfo10" name="txtInfo10" maxlength="15" size="15" /></td>
                      </tr>
                    </table></td>
                  <td style="border-left:solid; border-left-color:#CCC; border-left-width:1px; padding-left:10px;"><table width="400">
                      <tr>
                        <td>&nbsp;</td>
                        <td>Product</td>
                        <td>Qty</td>
                        <td>Pallet(s)</td>
                        <td>Info</td>
                      </tr>
                      <tr>
                        <td>11:</td>
                        <td><input type="text" id="txtProduct11" name="txtProduct11" maxlength="15" size="15" /></td>
                        <td><input type="text" id="txtQty11" name="txtQty11" maxlength="4" size="5" /></td>
                        <td><input type="text" id="txtPallet11" name="txtPallet11" maxlength="4" size="5" /></td>
                        <td><input type="text" id="txtInfo11" name="txtInfo11" maxlength="15" size="15" /></td>
                      </tr>
                      <tr>
                        <td>12:</td>
                        <td><input type="text" id="txtProduct12" name="txtProduct12" maxlength="15" size="15" /></td>
                        <td><input type="text" id="txtQty12" name="txtQty12" maxlength="4" size="5" /></td>
                        <td><input type="text" id="txtPallet12" name="txtPallet12" maxlength="4" size="5" /></td>
                        <td><input type="text" id="txtInfo12" name="txtInfo12" maxlength="15" size="15" /></td>
                      </tr>
                      <tr>
                        <td>13:</td>
                        <td><input type="text" id="txtProduct13" name="txtProduct13" maxlength="15" size="15" /></td>
                        <td><input type="text" id="txtQty13" name="txtQty13" maxlength="4" size="5" /></td>
                        <td><input type="text" id="txtPallet13" name="txtPallet13" maxlength="4" size="5" /></td>
                        <td><input type="text" id="txtInfo13" name="txtInfo13" maxlength="15" size="15" /></td>
                      </tr>
                      <tr>
                        <td>14:</td>
                        <td><input type="text" id="txtProduct14" name="txtProduct14" maxlength="15" size="15" /></td>
                        <td><input type="text" id="txtQty14" name="txtQty14" maxlength="4" size="5" /></td>
                        <td><input type="text" id="txtPallet14" name="txtPallet14" maxlength="4" size="5" /></td>
                        <td><input type="text" id="txtInfo14" name="txtInfo14" maxlength="15" size="15" /></td>
                      </tr>
                      <tr>
                        <td>15:</td>
                        <td><input type="text" id="txtProduct15" name="txtProduct15" maxlength="15" size="15" /></td>
                        <td><input type="text" id="txtQty15" name="txtQty15" maxlength="4" size="5" /></td>
                        <td><input type="text" id="txtPallet15" name="txtPallet15" maxlength="4" size="5" /></td>
                        <td><input type="text" id="txtInfo15" name="txtInfo15" maxlength="15" size="15" /></td>
                      </tr>
                      <tr>
                        <td>16:</td>
                        <td><input type="text" id="txtProduct16" name="txtProduct16" maxlength="15" size="15" /></td>
                        <td><input type="text" id="txtQty16" name="txtQty16" maxlength="4" size="5" /></td>
                        <td><input type="text" id="txtPallet16" name="txtPallet16" maxlength="4" size="5" /></td>
                        <td><input type="text" id="txtInfo16" name="txtInfo16" maxlength="15" size="15" /></td>
                      </tr>
                      <tr>
                        <td>17:</td>
                        <td><input type="text" id="txtProduct17" name="txtProduct17" maxlength="15" size="15" /></td>
                        <td><input type="text" id="txtQty17" name="txtQty17" maxlength="4" size="5" /></td>
                        <td><input type="text" id="txtPallet17" name="txtPallet17" maxlength="4" size="5" /></td>
                        <td><input type="text" id="txtInfo17" name="txtInfo17" maxlength="15" size="15" /></td>
                      </tr>
                      <tr>
                        <td>18:</td>
                        <td><input type="text" id="txtProduct18" name="txtProduct18" maxlength="15" size="15" /></td>
                        <td><input type="text" id="txtQty18" name="txtQty18" maxlength="4" size="5" /></td>
                        <td><input type="text" id="txtPallet18" name="txtPallet18" maxlength="4" size="5" /></td>
                        <td><input type="text" id="txtInfo18" name="txtInfo18" maxlength="15" size="15" /></td>
                      </tr>
                      <tr>
                        <td>19:</td>
                        <td><input type="text" id="txtProduct19" name="txtProduct19" maxlength="15" size="15" /></td>
                        <td><input type="text" id="txtQty19" name="txtQty19" maxlength="4" size="5" /></td>
                        <td><input type="text" id="txtPallet19" name="txtPallet19" maxlength="4" size="5" /></td>
                        <td><input type="text" id="txtInfo19" name="txtInfo19" maxlength="15" size="15" /></td>
                      </tr>
                      <tr>
                        <td>20:</td>
                        <td><input type="text" id="txtProduct20" name="txtProduct20" maxlength="15" size="15" /></td>
                        <td><input type="text" id="txtQty20" name="txtQty20" maxlength="4" size="5" /></td>
                        <td><input type="text" id="txtPallet20" name="txtPallet20" maxlength="4" size="5" /></td>
                        <td><input type="text" id="txtInfo20" name="txtInfo20" maxlength="15" size="15" /></td>
                      </tr>
                    </table></td>
                </tr>
              </table></td>
          </tr>
          <tr>
            <td>Delivery date<span class="mandatory">*</span>:</td>
            <td><input type="text" id="txtDeliveryDate" name="txtDeliveryDate" maxlength="10" size="8" />
              <em>DD/MM/YYYY</em></td>
          </tr>
          <tr>
            <td>Delivery time<span class="mandatory">*</span>:</td>
            <td><input type="text" id="txtDeliveryTime" name="txtDeliveryTime" maxlength="6" size="8" /></td>
          </tr>
          <tr>
            <td>Invoice no:</td>
            <td><input type="text" id="txtInvoiceNo" name="txtInvoiceNo" maxlength="20" size="30" /></td>
          </tr>
          <tr>
            <td valign="top">Comments:</td>
            <td><textarea name="txtComments" id="txtComments" cols="40" rows="4"></textarea></td>
          </tr>
        </table>
        <p><input type="hidden" name="Action" />
        <input type="submit" value="Add Transfer" />
      <input type="reset" value="Reset" /></p></td>
    </tr>
  </table>
</form>
<script type="text/javascript" src="include/moment.js"></script>
<script type="text/javascript" src="include/pikaday.js"></script>
<script type="text/javascript">
    var picker = new Pikaday(
    {
        field: document.getElementById('txtDeliveryDate'),
        firstDay: 1,
        minDate: new Date('2013-01-01'),
        maxDate: new Date('2030-12-31'),
        yearRange: [2013,2030],
        format: 'DD/MM/YYYY'
    });
</script>
</body>
</html>