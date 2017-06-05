<%
'----------------------------------------------------------------------------------------
' ADD PALLET
'----------------------------------------------------------------------------------------
Function addPallet(strPalletDepartment, strPalletInfo)
    dim strSQL

    call OpenDataBase()

    strSQL = "INSERT INTO tbl_pallets ("
    strSQL = strSQL & " pallet_department, "
    strSQL = strSQL & " pallet_info, "
    strSQL = strSQL & " pallet_created_by"
    strSQL = strSQL & " ) VALUES ( "
    strSQL = strSQL & "'" & strPalletDepartment & "',"
    strSQL = strSQL & "'" & Server.HTMLEncode(strPalletInfo) & "',"
    strSQL = strSQL & "'" & session("UsrUserName") & "')"

    'response.Write strSQL	

    on error resume next
    conn.Execute strSQL

    On error Goto 0

    if err <> 0 then
        strMessageText = err.description
    else
        strMessageText = "Pallet has been successfully added."
    end if

    call CloseDataBase()
end function

'----------------------------------------------------------------------------------------
' UPDATE PALLET
'----------------------------------------------------------------------------------------
Function updatePallet(intID, strPalletDepartment, strPalletInfo, intPalletStatus)
    dim strSQL

    Call OpenDataBase()

    strSQL = "UPDATE tbl_pallets SET "
    strSQL = strSQL & "pallet_department = '" & Server.HTMLEncode(strPalletDepartment) & "',"
    strSQL = strSQL & "pallet_info = '" & Server.HTMLEncode(strPalletInfo) & "',"
    strSQL = strSQL & "pallet_status = '" & intPalletStatus & "',"
    strSQL = strSQL & "pallet_date_modified = getdate(),"
    strSQL = strSQL & "pallet_modified_by = '" & session("UsrUserName") & "' WHERE pallet_id = " & intID

    'response.Write strSQL
    on error resume next
    conn.Execute strSQL

    On error Goto 0

    if err <> 0 then
        strMessageText = err.description
    else
        strMessageText = "The record has been updated."
    end if

    Call CloseDataBase()
end function

'----------------------------------------------------------------------------------------
' GET PALLET LIST
'----------------------------------------------------------------------------------------
function getPalletList
    dim strSQL
    dim rs
    dim intPalletID

    call OpenDataBase()

    strSQL = "SELECT * FROM tbl_pallets WHERE pallet_status = 1 ORDER BY pallet_id"

    set rs = server.CreateObject("ADODB.Recordset")
    set rs = conn.execute(strSQL)

    strPalletList = strPalletList & "<option value=''>...</option>"

    if not DB_RecSetIsEmpty(rs) Then
        do until rs.EOF
            intPalletID         = trim(rs("pallet_id"))
            strPalletDepartment = trim(rs("pallet_department"))

            if trim(session("report_pallet_no")) = intPalletID then
                strPalletList = strPalletList & "<option selected value=" & intPalletID & ">" & intPalletID & " (" & strPalletDepartment & ")</option>"
            else
                strPalletList = strPalletList & "<option value=" & intPalletID & ">" & intPalletID & " (" & strPalletDepartment & ")</option>"
            end if
        rs.Movenext
        loop
    end if

    call CloseDataBase()
end function

'----------------------------------------------------------------------------------------
' GET ALL PALLET LIST
'----------------------------------------------------------------------------------------
function getAllPalletList
    dim strSQL
    dim rs
    dim intPalletID
    dim strPalletDepartment

    call OpenDataBase()

    strSQL = "SELECT * FROM tbl_pallets WHERE pallet_status = 1 ORDER BY pallet_id"

    set rs = server.CreateObject("ADODB.Recordset")
    set rs = conn.execute(strSQL)

    strAllPalletList = strAllPalletList & "<option value=''>All Pallets</option>"

    if not DB_RecSetIsEmpty(rs) Then
        do until rs.EOF
            intPalletID         = trim(rs("pallet_id"))
            strPalletDepartment = trim(rs("pallet_department"))

            if trim(session("gra_report_pallet")) = intPalletID then
                strAllPalletList = strAllPalletList & "<option selected value=" & intPalletID & ">" & intPalletID & " (" & strPalletDepartment & ")</option>"
            else
                strAllPalletList = strAllPalletList & "<option value=" & intPalletID & ">" & intPalletID & " (" & strPalletDepartment & ")</option>"
            end if
        rs.Movenext
        loop
    end if

    call CloseDataBase()
end function

'----------------------------------------------------------------------------------------
' GET PALLET DETAILS
'----------------------------------------------------------------------------------------
function getPalletDetails(intPalletID)
    dim strSQL
    dim rs

    call OpenDataBase()

    set rs = Server.CreateObject("ADODB.recordset")

    rs.CursorLocation = 3   'adUseClient
    rs.CursorType = 3       'adOpenStatic

    strSQL = "SELECT * FROM tbl_pallets WHERE pallet_id = " & intPalletID

    rs.Open strSQL, conn

    if not DB_RecSetIsEmpty(rs) Then
        session("pallet_type")          = trim(rs("pallet_type"))
        session("pallet_department")    = trim(rs("pallet_department"))
        session("pallet_info")          = trim(rs("pallet_info"))
        session("pallet_date_created")  = trim(rs("pallet_date_created"))
        session("pallet_created_by")    = trim(rs("pallet_created_by"))
        session("pallet_date_modified") = trim(rs("pallet_date_modified"))
        session("pallet_modified_by")   = trim(rs("pallet_modified_by"))
    end if

    call CloseDataBase()
end function
%>