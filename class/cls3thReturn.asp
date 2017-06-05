<%
'-----------------------------------------------
' ADD 3TH RETURN
'-----------------------------------------------
function add3thReturn(intReturnType, strDepartment, strItemCode, strShipmentNo, intQty, strDescription, strGRA, strCarrier, strLabelNo, strOriginalConnote, strDealer, intInstruction, strSerialNo, intStockType, strDateReceived, strComments, strCreatedBy)
    dim strSQL

    Call OpenDataBase()

    strSQL = "INSERT INTO yma_3th_return ("
    strSQL = strSQL & " return_type, "
    strSQL = strSQL & " department, "
    strSQL = strSQL & " item_code, "
    strSQL = strSQL & " shipment_no, "
    strSQL = strSQL & " qty, "
    strSQL = strSQL & " description, "
    strSQL = strSQL & " gra, "
    strSQL = strSQL & " carrier, "
    strSQL = strSQL & " label_no, "
    strSQL = strSQL & " original_connote, "
    strSQL = strSQL & " dealer, "
    strSQL = strSQL & " instruction, "
    strSQL = strSQL & " serial_no, "
    strSQL = strSQL & " stock_type, "
    strSQL = strSQL & " date_received, "
    strSQL = strSQL & " comments, "
    strSQL = strSQL & " created_by"
    strSQL = strSQL & ") VALUES ("
    strSQL = strSQL & "'" & intReturnType & "',"
    strSQL = strSQL & "'" & strDepartment & "',"
    strSQL = strSQL & "'" & Server.HTMLEncode(strItemCode) & "',"
    strSQL = strSQL & "'" & Server.HTMLEncode(strShipmentNo) & "',"
    strSQL = strSQL & "'" & intQty & "',"
    strSQL = strSQL & "'" & Server.HTMLEncode(strDescription) & "',"
    strSQL = strSQL & "'" & Server.HTMLEncode(strGRA) & "',"
    strSQL = strSQL & "'" & strCarrier & "',"
    strSQL = strSQL & "'" & Server.HTMLEncode(strLabelNo) & "',"
    strSQL = strSQL & "'" & Server.HTMLEncode(strOriginalConnote) & "',"
    strSQL = strSQL & "'" & Server.HTMLEncode(strDealer) & "',"	
    strSQL = strSQL & "'" & intInstruction & "',"
    strSQL = strSQL & "'" & Server.HTMLEncode(strSerialNo) & "',"
    strSQL = strSQL & "'" & intStockType & "',"
    strSQL = strSQL & " CONVERT(DateTime,'" & strDateReceived & "',103),"
    strSQL = strSQL & "'" & Server.HTMLEncode(strComments) & "',"
    strSQL = strSQL & "'" & strCreatedBy & "')"

    'response.Write strSQL

    on error resume next
    conn.Execute strSQL

    if err <> 0 then
        strMessageText = err.description
    else
        strMessageText = "<div class=""notification_text""><img src=""images/icon_check.png""> 3TH Return has been successfully added.</div>"
    end if

    Call CloseDataBase()
end function

'-----------------------------------------------
' GET 3TH RETURN
'-----------------------------------------------
Function get3thReturn(intReturnID)
    dim strTodayDate
    strTodayDate = FormatDateTime(Date())

    call OpenDataBase()

    set rs = Server.CreateObject("ADODB.recordset")

    rs.CursorLocation = 3   'adUseClient
    rs.CursorType = 3       'adOpenStatic

    strSQL = "SELECT *"
    strSQL = strSQL & "	FROM yma_3th_return"
    strSQL = strSQL & "	WHERE return_id = '" & intReturnID & "'"

    rs.Open strSQL, conn

    'response.write strSQL

    if not DB_RecSetIsEmpty(rs) Then
        session("return_type")      = rs("return_type")
        session("3thDepartment")    = rs("department")
        session("item_code")        = rs("item_code")
        session("shipment_no")      = rs("shipment_no")
        session("qty")              = rs("qty")
        session("description")      = rs("description")
        session("gra")              = rs("gra")
        session("carrier")          = rs("carrier")
        session("label_no")         = rs("label_no")
        session("original_connote") = rs("original_connote")
        session("dealer")           = rs("dealer")
        session("reason_code")      = rs("reason_code")
        session("instruction")      = rs("instruction")
        session("serial_no")        = rs("serial_no")
        session("stock_type")       = rs("stock_type")
        session("date_received")    = rs("date_received")
        session("comments")         = rs("comments")
        session("date_created")     = rs("date_created")
        session("created_by")       = rs("created_by")
        session("date_modified")    = rs("date_modified")
        session("modified_by")      = rs("modified_by")
        session("status")           = rs("status")
        session("days_in_3TH")      = DateDiff("d",rs("date_created"), strTodayDate)
    end if

    call CloseDataBase()
end Function

'-----------------------------------------------
' UPDATE 3TH WAREHOUSE
'-----------------------------------------------
Function update3thReturn(intReturnID, intReturnType, strDepartment, strItemCode, strShipmentNo, intQty, strDescription, strGRA, strCarrier, strLabelNo, strOriginalConnote, strDealer, intInstruction, strSerialNo, strDateReceived, strComments, intStatus, strModifiedBy)	
    dim strSQL

    Call OpenDataBase()

    strSQL = "UPDATE yma_3th_return SET "
    strSQL = strSQL & " return_type = '" & intReturnType & "',"
    strSQL = strSQL & " department = '" & strDepartment & "',"
    strSQL = strSQL & " item_code = '" & Server.HTMLEncode(strItemCode) & "',"
    strSQL = strSQL & " shipment_no = '" & Server.HTMLEncode(strShipmentNo) & "',"
    strSQL = strSQL & " qty = '" & intQty & "',"
    strSQL = strSQL & " description = '" & Server.HTMLEncode(strDescription) & "',"
    strSQL = strSQL & " gra = '" & Server.HTMLEncode(strGRA) & "',"
    strSQL = strSQL & " carrier = '" & strCarrier & "',"
    strSQL = strSQL & " label_no = '" & Server.HTMLEncode(strLabelNo) & "',"
    strSQL = strSQL & " original_connote = '" & Server.HTMLEncode(strOriginalConnote) & "',"
    strSQL = strSQL & " dealer = '" & Server.HTMLEncode(strDealer) & "',"
    strSQL = strSQL & " instruction = '" & intInstruction & "',"
    strSQL = strSQL & " serial_no = '" & Server.HTMLEncode(strSerialNo) & "',"
    strSQL = strSQL & " date_received = CONVERT(datetime,'" & strDateReceived & "',103),"
    strSQL = strSQL & " comments = '" & Server.HTMLEncode(strComments) & "',"
    strSQL = strSQL & " status = '" & intStatus & "',"
    strSQL = strSQL & " date_modified = GetDate(),"
    strSQL = strSQL & " modified_by = '" & strModifiedBy & "' "
    strSQL = strSQL & " 	WHERE return_id = " & intReturnID

    'response.Write strSQL
    on error resume next
    conn.Execute strSQL

    if err <> 0 then
        strMessageText = err.description
    else
        strMessageText = "<div class=""notification_text""><img src=""images/icon_check.png""> 3TH Return has been updated.</div>"
    end if

    Call CloseDataBase()
end Function
%>