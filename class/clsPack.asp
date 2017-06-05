<%
'-----------------------------------------------
' LOGISTICS CONFIRM PACK
'-----------------------------------------------
Function logisticsConfirmPack(packID, packName, packQty, packName2, packQty2, packName3, packQty3, packName4, packQty4, packName5, packQty5, packPriority, packModifiedBy)
    dim strSQL

    Call OpenDataBase()

    strSQL = "UPDATE logistic_pack SET "
    strSQL = strSQL & "packLogistics = 1,"
    strSQL = strSQL & "packLogisticsDate = GetDate(),"
    strSQL = strSQL & "packDateModified = GetDate(),"
    strSQL = strSQL & "packModifiedBy = '" & packModifiedBy & "'"
    strSQL = strSQL & " WHERE packID = " & packID

    'response.Write strSQL

    on error resume next
    conn.Execute strSQL

    if err <> 0 then
        strMessageText = err.description
    else
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

        emailFrom       = "noreply@music.yamaha.com"
        'emailTo         = "harsono_setiono@gmx.yamaha.com"
        emailTo         = "YMA-Warehouse@ttlogistics.com.au"
        emailCc         = "logistics-aus@music.yamaha.com"
        emailSubject    = "New Pack (Confirmed by Logistics)"
        emailBodyText   = "G'day Warehouse," & vbCrLf _
                        & " " & vbCrLf _
                        & "This following new pack has been confirmed by Logistics and needs to be processed." & vbCrLf _
                        & "---------------------------------------------------------------------------" & vbCrLf _
                        & "1. Name : " & packName & vbCrLf _
                        & "   Qty  : " & packQty & vbCrLf _
                        & "2. Name : " & packName2 & vbCrLf _
                        & "   Qty  : " & packQty2 & vbCrLf _
                        & "3. Name : " & packName3 & vbCrLf _
                        & "   Qty  : " & packQty3 & vbCrLf _
                        & "4. Name : " & packName4 & vbCrLf _
                        & "   Qty  : " & packQty4 & vbCrLf _
                        & "5. Name : " & packName5 & vbCrLf _
                        & "   Qty  : " & packQty5 & vbCrLf _
                        & " " & vbCrLf _
                        & "Priority : " & packPriority & vbCrLf _
                        & "---------------------------------------------------------------------------" & vbCrLf _
                        & "Please click on the below link to process it:" & vbCrLf _
                        & "http://intranet/logistics/list_pack.asp" & vbCrLf _
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

        Response.Redirect(Request.ServerVariables("HTTP_REFERER"))
    end if

    Call CloseDataBase()
end Function

'-----------------------------------------------
' WAREHOUSE CONFIRM PACK
'-----------------------------------------------
Function warehouseConfirmPack(packID, packName, packQty, packName2, packQty2, packName3, packQty3, packName4, packQty4, packName5, packQty5, packPriority, packModifiedBy, packEmail)
    dim strSQL

    Call OpenDataBase()

    strSQL = "UPDATE logistic_pack SET "
    strSQL = strSQL & "packStatus = 0,"
    strSQL = strSQL & "packWarehouse = 1,"
    strSQL = strSQL & "packWarehouseDate = GetDate(),"
    strSQL = strSQL & "packDateModified = GetDate(),"
    strSQL = strSQL & "packModifiedBy = '" & packModifiedBy & "'"
    strSQL = strSQL & " WHERE packID = " & packID

    'response.Write strSQL

    on error resume next
    conn.Execute strSQL

    if err <> 0 then
        strMessageText = err.description
    else
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

        emailFrom       = "noreply@music.yamaha.com"
        emailTo         = packEmail
        emailCc         = "logistics-aus@music.yamaha.com"
        emailSubject    = "Your Pack Request has been processed"
        emailBodyText   = "G'day!" & vbCrLf _
                        & " " & vbCrLf _
                        & "The new pack you requested has been processed by Logistics and Warehouse." & vbCrLf _
                        & "---------------------------------------------------------------------------" & vbCrLf _
                        & "1. Name : " & packName & vbCrLf _
                        & "   Qty  : " & packQty & vbCrLf _
                        & "2. Name : " & packName2 & vbCrLf _
                        & "   Qty  : " & packQty2 & vbCrLf _
                        & "3. Name : " & packName3 & vbCrLf _
                        & "   Qty  : " & packQty3 & vbCrLf _
                        & "4. Name : " & packName4 & vbCrLf _
                        & "   Qty  : " & packQty4 & vbCrLf _
                        & "5. Name : " & packName5 & vbCrLf _
                        & "   Qty  : " & packQty5 & vbCrLf _
                        & " " & vbCrLf _
                        & "Priority : " & packPriority & vbCrLf _
                        & "---------------------------------------------------------------------------" & vbCrLf _
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

        Response.Redirect(Request.ServerVariables("HTTP_REFERER"))
    end if

    Call CloseDataBase()
end Function

'-----------------------------------------------
' WAREHOUSE SET ETA
' CREATED: Victor 2016-07-18
'-----------------------------------------------
Function warehouseSetETA(packID, packWarehouseETA, packModifiedBy)
    Dim strSQL

    Call OpenDataBase()

    strSQL          = "UPDATE dbo.logistic_pack SET "
    strSQL = strSQL & "packWarehouseETA = CONVERT(datetime, '" & packWarehouseETA & "', 103), "
    strSQL = strSQL & "packDateModified = GetDate(), "
    strSQL = strSQL & "packModifiedBy = '" & packModifiedBy & "' "
    strSQL = strSQL & "WHERE packID = " & packID

    'Response.Write strSQL
    on error resume next
    conn.Execute strSQL

    If err <> 0 Then
        strMessageText = err.description
    Else
        Response.Redirect(Request.ServerVariables("HTTP_REFERER"))
    End If

    Call CloseDataBase()
End Function

'-----------------------------------------------
' UPDATE PACK COMMENTS
'-----------------------------------------------
Function updatePackComments(packID, packComments, packModifiedBy)
    dim strSQL

    Call OpenDataBase()

    strSQL = "UPDATE logistic_pack SET "
    strSQL = strSQL & "packComments = '" & Server.HTMLEncode(packComments) & "',"
    strSQL = strSQL & "packDateModified = GetDate(),"
    strSQL = strSQL & "packModifiedBy = '" & packModifiedBy & "'"
    strSQL = strSQL & " WHERE packID = " & packID

    'response.Write strSQL
    on error resume next
    conn.Execute strSQL

    if err <> 0 then
        strMessageText = err.description
    else
        Response.Redirect(Request.ServerVariables("HTTP_REFERER"))
    end if

    Call CloseDataBase()
end Function
%>