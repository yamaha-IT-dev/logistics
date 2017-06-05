<%
'-----------------------------------------------
' GET SALES GROUP
'-----------------------------------------------
' Used in:  add_item-maintenance.asp
'           update_item-maintenance.asp
'-----------------------------------------------
function getSalesGroupList
    dim strSQL
    dim rs
    dim intGroupID
    dim strGroupName

    call OpenBaseDataBase()

    strSQL = "SELECT YDSGID, YDSGCD, YDSGMB FROM YFDMP WHERE YDSGID = '3' AND YDSKKI <> 'D' ORDER BY YDSGCD"

    set rs = server.CreateObject("ADODB.Recordset")
    set rs = conn.execute(strSQL)

    strSalesGroupList = strSalesGroupList & "<option value=''>...</option>"

    if not DB_RecSetIsEmpty(rs) Then
        do until rs.EOF
            intGroupID      = trim(rs("YDSGCD"))
            strGroupName    = UCase(Left(rs("YDSGMB"), 1))
            strGroupName    = strGroupName & LCase(Right(rs("YDSGMB"), len(rs("YDSGMB")) - 1))

            if trim(session("sales_group")) = intGroupID then
                strSalesGroupList = strSalesGroupList & "<option selected value=" & intGroupID & ">" & intGroupID & ". " & strGroupName & "</option>"
            else
                strSalesGroupList = strSalesGroupList & "<option value=" & intGroupID & ">" & intGroupID & ". " & strGroupName & "</option>"
            end if

        rs.Movenext
        loop

    end if

    call CloseDataBase()
end function

'-----------------------------------------------
' GET SHIPPING GROUP
'-----------------------------------------------
' Used in:  add_item-maintenance.asp
'           update_item-maintenance.asp
'-----------------------------------------------
function getShippingGroupList
    dim strSQL
    dim rs
    dim intGroupID
    dim strGroupName

    call OpenBaseDataBase()

    strSQL = "SELECT YDSGID, YDSGCD, YDSGMB FROM YFDMP WHERE YDSGID = '2' AND YDSKKI <> 'D' AND YDSGCD <> '0' ORDER BY YDSGCD"

    set rs = server.CreateObject("ADODB.Recordset")
    set rs = conn.execute(strSQL)

    strShippingGroupList = strShippingGroupList & "<option value=''>...</option>"

    if not DB_RecSetIsEmpty(rs) Then
        do until rs.EOF
            intGroupID      = trim(rs("YDSGCD"))
            strGroupName    = UCase(Left(rs("YDSGMB"), 1))
            strGroupName    = strGroupName & LCase(Right(rs("YDSGMB"), len(rs("YDSGMB")) - 1))

            if trim(session("shipping_group")) = intGroupID then
                strShippingGroupList = strShippingGroupList & "<option selected value=" & intGroupID & ">" & intGroupID & ". " & strGroupName & "</option>"
            else
                strShippingGroupList = strShippingGroupList & "<option value=" & intGroupID & ">" & intGroupID & ". " & strGroupName & "</option>"
            end if

        rs.Movenext
        loop

    end if

    call CloseDataBase()
end function

'-----------------------------------------------
' GET ACCOUNT GROUP
'-----------------------------------------------
' Used in:  add_item-maintenance.asp
'           update_item-maintenance.asp
'-----------------------------------------------
function getAccountGroupList
    dim strSQL
    dim rs
    dim intGroupID
    dim strGroupName

    call OpenBaseDataBase()

    strSQL = "SELECT YDSGID, YDSGCD, YDSGMB FROM YFDMP WHERE YDSGID = '1' AND YDSKKI <> 'D' ORDER BY YDSGCD"

    set rs = server.CreateObject("ADODB.Recordset")
    set rs = conn.execute(strSQL)

    strAccountGroupList = strAccountGroupList & "<option value=''>...</option>"

    if not DB_RecSetIsEmpty(rs) Then
        do until rs.EOF
            intGroupID      = trim(rs("YDSGCD"))
            strGroupName    = UCase(Left(rs("YDSGMB"), 1))
            strGroupName    = strGroupName & LCase(Right(rs("YDSGMB"), len(rs("YDSGMB")) - 1))

            if trim(session("account_group")) = intGroupID then
                strAccountGroupList = strAccountGroupList & "<option selected value=" & intGroupID & ">" & intGroupID & ". " & strGroupName & "</option>"
            else
                strAccountGroupList = strAccountGroupList & "<option value=" & intGroupID & ">" & intGroupID & ". " & strGroupName & "</option>"
            end if

        rs.Movenext
        loop

    end if

    call CloseDataBase()
end function

'-----------------------------------------------
' GET DISCOUNT GROUP
'-----------------------------------------------
' Used in:  add_item-maintenance.asp
'           update_item-maintenance.asp
'-----------------------------------------------
function getDiscountGroupList
    dim strSQL
    dim rs
    dim strDiscountID

    call OpenBaseDataBase()

    'strSQL = "SELECT DISTINCT YBGRDS FROM YFBMP WHERE YBSKKI <> 'D' AND YBGRDS <> '' ORDER BY YBGRDS"
    strSQL = "SELECT YCKBCD FROM YFCMP WHERE YCKBID = 'GRDS' AND YCSKKI <> 'D' ORDER BY YCKBCD"
    set rs = server.CreateObject("ADODB.Recordset")
    set rs = conn.execute(strSQL)

    strDiscountGroupList = strDiscountGroupList & "<option value=''>...</option>"

    if not DB_RecSetIsEmpty(rs) Then
        do until rs.EOF
            strDiscountID   = trim(rs("YCKBCD"))

            if trim(session("discount_group")) = strDiscountID then
                strDiscountGroupList = strDiscountGroupList & "<option selected value=" & strDiscountID & ">" & strDiscountID & "</option>"
            else
                strDiscountGroupList = strDiscountGroupList & "<option value=" & strDiscountID & ">" & strDiscountID & "</option>"
            end if

        rs.Movenext
        loop

    end if

    call CloseDataBase()
end function

'-----------------------------------------------
' GET TARIFF CODE
'-----------------------------------------------
' Used in:  add_item-maintenance.asp
'           update_item-maintenance.asp
'-----------------------------------------------
function getTariffCodeList
    dim strSQL
    dim rs
    dim strTariffCode

    call OpenBaseDataBase()

    strSQL = "SELECT DISTINCT W9SKKI, W9KZEC FROM WF9MP WHERE W9SKKI <> 'D' ORDER BY W9KZEC"

    set rs = server.CreateObject("ADODB.Recordset")
    set rs = conn.execute(strSQL)

    strTariffCodeList = strTariffCodeList & "<option value=''>...</option>"

    if not DB_RecSetIsEmpty(rs) Then
        do until rs.EOF
            strTariffCode   = trim(rs("W9KZEC"))

            if trim(session("tariff_code")) = strTariffCode then
                strTariffCodeList = strTariffCodeList & "<option selected value=" & strTariffCode & ">" & strTariffCode & "</option>"
            else
                strTariffCodeList = strTariffCodeList & "<option value=" & strTariffCode & ">" & strTariffCode & "</option>"
            end if

        rs.Movenext
        loop

    end if

    call CloseDataBase()
end function

'-----------------------------------------------
' GET COUNTRIES
'-----------------------------------------------
' Used in:  add_item-maintenance.asp
'           update_item-maintenance.asp
'-----------------------------------------------
function getCountryList
    dim strSQL
    dim rs
    dim strCountryID
    dim strCountryName

    call OpenDataBase()

    strSQL = "SELECT * FROM tbl_countries ORDER BY country_name"

    set rs = server.CreateObject("ADODB.Recordset")
    set rs = conn.execute(strSQL)

    strCountryList = strCountryList & "<option value=''>...</option>"

    if not DB_RecSetIsEmpty(rs) Then
        do until rs.EOF
            strCountryID    = trim(rs("country_id"))
            strCountryName  = trim(rs("country_name"))

            if trim(session("country_origin")) = strCountryID then
                strCountryList = strCountryList & "<option selected value=" & strCountryID & ">" & strCountryName & "</option>"
            else
                strCountryList = strCountryList & "<option value=" & strCountryID & ">" & strCountryName & "</option>"
            end if

        rs.Movenext
        loop

    end if

    call CloseDataBase()
end function

'-----------------------------------------------
' GET VENDORS
'-----------------------------------------------
' Used in:  add_item-maintenance.asp
'           update_item-maintenance.asp
'-----------------------------------------------
function getVendorList
    dim strSQL
    dim rs
    dim strVendorID
    dim strVendorName

    call OpenDataBase()

    strSQL = "SELECT * FROM tbl_vendors ORDER BY vendor_name"

    set rs = server.CreateObject("ADODB.Recordset")
    set rs = conn.execute(strSQL)

    strVendorList = strVendorList & "<option value=''>...</option>"

    if not DB_RecSetIsEmpty(rs) Then
        do until rs.EOF
            strVendorID     = trim(rs("vendor_id"))
            strVendorName   = trim(rs("vendor_name"))

            if trim(session("vendor")) = strVendorID then
                strVendorList = strVendorList & "<option selected value=" & strVendorID & ">" & strVendorName & "</option>"
            else
                strVendorList = strVendorList & "<option value=" & strVendorID & ">" & strVendorName & "</option>"
            end if

        rs.Movenext
        loop

    end if

    call CloseDataBase()
end function

'-----------------------------------------------
' GET LIFECYCLES
'-----------------------------------------------
' Used in:  add_item-maintenance.asp
'           update_item-maintenance.asp
'-----------------------------------------------
function getLifecycleList
    dim strSQL
    dim rs
    dim strLifecycleID
    dim strLifecycleName

    call OpenDataBase()

    strSQL = "SELECT * FROM tbl_lifecycles"

    set rs = server.CreateObject("ADODB.Recordset")
    set rs = conn.execute(strSQL)

    strLifecycleList = strLifecycleList & "<option value=''>...</option>"

    if not DB_RecSetIsEmpty(rs) Then
        do until rs.EOF
            strLifecycleID      = trim(rs("lifecycle_id"))
            strLifecycleName    = trim(rs("lifecycle_name"))

            if trim(session("lifecycle")) = strLifecycleID then
                strLifecycleList = strLifecycleList & "<option selected value=" & strLifecycleID & ">" & strLifecycleID & ": " & strLifecycleName & "</option>"
            else
                strLifecycleList = strLifecycleList & "<option value=" & strLifecycleID & ""
                if strLifecycleID = "D" then
                    strLifecycleList = strLifecycleList & " rel=""discontinued"">"
                else
                    strLifecycleList = strLifecycleList & " rel=""none"">"
                end if
                strLifecycleList = strLifecycleList & "" & strLifecycleID & ": " & strLifecycleName & "</option>"
            end if

        rs.Movenext
        loop

    end if

    call CloseDataBase()
end function

'-----------------------------------------------
' GET COLOURS
'-----------------------------------------------
' Used in:  add_item-maintenance.asp
'           update_item-maintenance.asp
'-----------------------------------------------
function getColourList
    dim strSQL
    dim rs
    dim strColourID
    dim strColourName

    call OpenDataBase()

    strSQL = "SELECT * FROM tbl_colours ORDER BY colour_id"

    set rs = server.CreateObject("ADODB.Recordset")
    set rs = conn.execute(strSQL)

    strColourList = strColourList & "<option value=''>...</option>"

    if not DB_RecSetIsEmpty(rs) Then
        do until rs.EOF
            strColourID    = trim(rs("colour_id"))
            strColourName  = trim(rs("colour_name"))

            if trim(session("colour")) = strColourID then
                strColourList = strColourList & "<option selected value=" & strColourID & ">" & strColourName & "</option>"
            else
                strColourList = strColourList & "<option value=" & strColourID & ">" & strColourName & "</option>"
            end if

        rs.Movenext
        loop

    end if

    call CloseDataBase()
end function

'-------------------------------------------------------------
' APPROVE ITEM BY GENERAL MANAGER
' Used in:  update_item-maintenance.asp	
'-------------------------------------------------------------
function approveGM
    dim sql

    call OpenDataBase()

    sql = "UPDATE yma_item_maintenance SET gm_approval = 1, gm_approval_date = getDate(), date_modified = getDate() WHERE item_id = " & session("item_id")

    on error resume next
    conn.Execute sql

    On error Goto 0

    if err <> 0 then
        strMessageText = err.description
    else
        strMessageText = "GM Approved!"
    end if

    call CloseDataBase()
end function

'-------------------------------------------------------------
' FLAG ITEM MAINTENANCE AS CREATED IN BASE
' Used in: update_item-maintenance-base.asp 
'-------------------------------------------------------------
function itemCreatedInBase
    dim sql

    call OpenDataBase()

    sql = "UPDATE dbo.yma_item_maintenance SET logistics_pending = 1 WHERE item_id = " & session("item_id")

    on error resume next
    conn.Execute sql

    On error Goto 0

    if err <> 0 then
        itemCreatedInBase = err.description
    else
        itemCreatedInBase = "SUCCESS"
    end if

    call CloseDataBase()
end function

%>