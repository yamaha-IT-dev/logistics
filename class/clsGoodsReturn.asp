<%
'----------------------------------------------------------------------------------------
' Get GRA from BASE
'----------------------------------------------------------------------------------------
function getGraFromBASE(intID)
    dim strSQL

    call OpenBaseDataBase()

    set rs = Server.CreateObject("ADODB.recordset")

    rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic

    strSQL = "SELECT BTUSNO, BTKASC, BTOPEC, BTHYNO, BTURKC, BTHSRC, BTAHYY, BTAHYM, BTAHYD, BTAHYY, BTAHYM, BTAHYD, BTRTST, BTUSGR, BTATT1, BTSCNO, BTNICM, BTGICM, "
    strSQL = strSQL & " Y1KOM1, Y1JSO1, Y1JSO3, Y1KNCI, Y1UBNB, Y1TELN, YMOPEC, YMOPNM, YMSKKI, A3HSM1, A3HJS1, A3HJS3, A3KNCH, A3UBNH, A3HTLN "
    strSQL = strSQL & " 	FROM BFTEP "
    strSQL = strSQL & "			LEFT JOIN YF1MP ON CONCAT(BTURKC, BTHSRC) = Y1KOKC "
    strSQL = strSQL & "			LEFT JOIN AF3EP ON BTHYNO = A3HYNO "

    strSQL = strSQL & "			LEFT JOIN YFMMP ON BTOPEC = YMOPEC "
    'strSQL = strSQL & "			LEFT JOIN (SELECT YMOPEC, MAX(('20' + RIGHT(RTRIM(YMKOYY),2)) * 10000 + SUBSTRING(CONVERT(VARCHAR(6), YMKOYY), 3,2) * 100 + LEFT(YMKOYY,2)) AS MOD_DATE"
    'strSQL = strSQL & "			FROM YFMMP"
    'strSQL = strSQL & "			GROUP BY YMOPEC"
    'strSQL = strSQL & "			) AS OP ON BTOPEC = OP.YMOPEC"
    'strSQL = strSQL & "			INNER JOIN (SELECT YMOPEC, YMOPNM, YMPMID, (('20' + RIGHT(RTRIM(YMKOYY),2)) * 10000 + SUBSTRING(CONVERT(VARCHAR(6), YMKOYY), 3,2) * 100 + LEFT(YMKOYY,2)) AS MOD_DATE"
    'strSQL = strSQL & "				FROM YFMMP"
    'strSQL = strSQL & "			) AS OP_NAME ON OP.YMOPEC = OP_NAME.YMOPEC AND OP.MOD_DATE = OP_NAME.MOD_DATE"

    'strSQL = strSQL & "		WHERE BTSKKI <> 'D' " 'Removed 16 Oct 2014 due to Archived
    strSQL = strSQL & "		WHERE "
    strSQL = strSQL & "			(YMPMID like 'LOG%' or YMPMID like 'INT%' or YMPMID like 'OTH%' or YMPMID like 'SERV%' or YMPMID like 'CREDIT%' OR YMPMID like 'EXCE%') "
    'strSQL = strSQL & "			AND YMSKKI <> 'D' "
    'strSQL = strSQL & "			AND (YMSKKI <> 'D' OR (YMSKKI = 'D' AND YMOPEC = 'AC')) "
    strSQL = strSQL & "			AND BTHYNO = '" & intID & "'"
    'strSQL = strSQL & "			AND Y1SKKI <> 'D'"

    rs.Open strSQL, conn

    'Response.Write strSQL

    if not DB_RecSetIsEmpty(rs) Then
        session("gra_operator_code")        = trim(rs("BTOPEC"))
        session("gra_operator_name")        = trim(rs("YMOPNM"))
        session("gra_no")                   = trim(rs("BTHYNO"))
        session("gra_dealer_code")          = trim(rs("BTURKC"))
        session("gra_ship_to_dealer")       = trim(rs("BTHSRC"))
        if trim(rs("BTHSRC")) = "999" then
            session("gra_dealer_name")      = trim(rs("A3HSM1"))
            session("gra_dealer_address")   = trim(rs("A3HJS1"))
            session("gra_dealer_city")      = trim(rs("A3HJS3"))
            'session("gra_dealer_state")    = trim(rs("Y1KNCI"))
            Select Case trim(rs("A3KNCH"))
            case "01"
                session("gra_dealer_state") = "ACT"
            case "02"
                session("gra_dealer_state") = "NSW"
            case "03"
                session("gra_dealer_state") = "VIC"
            case "04"
                session("gra_dealer_state") = "QLD"
            case "05"
                session("gra_dealer_state") = "SA"
            case "06"
                session("gra_dealer_state") = "WA"
            case "07"
                session("gra_dealer_state") = "TAS"
            case "08"
                session("gra_dealer_state") = "NT"
            case else
                session("gra_dealer_state") = trim(rs("A3KNCH"))
            end select

            session("gra_dealer_postcode")  = trim(rs("A3UBNH"))
            session("gra_dealer_phone")     = trim(rs("A3HTLN"))

        else
            session("gra_dealer_name")      = trim(rs("Y1KOM1"))
            session("gra_dealer_address")   = trim(rs("Y1JSO1"))
            session("gra_dealer_city")      = trim(rs("Y1JSO3"))
            'session("gra_dealer_state")    = trim(rs("Y1KNCI"))
            Select Case trim(rs("Y1KNCI"))
            case "01"
                session("gra_dealer_state") = "ACT"
            case "02"
                session("gra_dealer_state") = "NSW"
            case "03"
                session("gra_dealer_state") = "VIC"
            case "04"
                session("gra_dealer_state") = "QLD"
            case "05"
                session("gra_dealer_state") = "SA"
            case "06"
                session("gra_dealer_state") = "WA"
            case "07"
                session("gra_dealer_state") = "TAS"
            case "08"
                session("gra_dealer_state") = "NT"
            case else
                session("gra_dealer_state") = trim(rs("Y1UBNB"))
            end select

            session("gra_dealer_postcode")  = trim(rs("Y1UBNB"))
            session("gra_dealer_phone")     = trim(rs("Y1TELN"))

        end if
        session("gra_day_entered")          = trim(rs("BTAHYD"))
        session("gra_month_entered")        = trim(rs("BTAHYM"))
        session("gra_year_entered")         = trim(rs("BTAHYY"))
        session("gra_return_status")        = trim(rs("BTRTST"))
        session("gra_carrier_code")         = trim(rs("BTUSGR"))
        session("gra_contact_person")       = trim(rs("BTATT1"))
        session("gra_warehouse")            = trim(rs("BTSCNO"))
        session("gra_ext_comment")          = trim(rs("BTNICM"))
        session("gra_int_comment")          = trim(rs("BTGICM"))
        session("gra_not_found")            = "FALSE"
    else
        session("gra_not_found")            = "TRUE"
        'response.redirect("gra-not-found.asp")
    end if

    call CloseBaseDataBase()
end function

'----------------------------------------------------------------------------------------
' Get GRA STATUS
'----------------------------------------------------------------------------------------
Function getGraStatus(intID)
    dim strSQL

    session("gra_status") = ""

    set rs = Server.CreateObject("ADODB.recordset")

    rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic

    Call OpenDataBase()

    strSQL = "SELECT * FROM yma_gra_status WHERE gra_no = '" & intID & "'"

    'Response.Write strSQL & "<br>"

    rs.Open strSQL, conn

    if not DB_RecSetIsEmpty(rs) Then
        session("gra_status") = trim(rs("status"))
    end if

    if session("gra_status") <> "" then
        select case session("gra_status")
            case 0
                session("gra_status_label") = "<font class=""yellow_font"">Need to sent to dealer</font>"
            case 1
                session("gra_status_label") = "<font class=""green_font"">Sent to dealer</font>"
        end select
    else
        session("gra_status_label") = "<font class=""blue_font"">New</font>"
    end if

    call CloseDataBase()
end function

'-----------------------------------------------
' ADD CONNOTE INTO yma_gra_status
'-----------------------------------------------
function addGraConnote(intGraNo,strConnote)
    dim strSQL

    Call OpenDataBase()

    strSQL = "INSERT INTO yma_gra_status ("
    strSQL = strSQL & "gra_no,"
    strSQL = strSQL & "gra_connote,"
    strSQL = strSQL & "gra_created_by"
    strSQL = strSQL & ") VALUES ("
    strSQL = strSQL & "'" & intGraNo & "',"
    strSQL = strSQL & "'" & Server.HTMLEncode(strConnote) & "',"
    strSQL = strSQL & "'" & session("UsrUserName") & "')"

    on error resume next
    conn.Execute strSQL

    'response.write strSQL

    if err <> 0 then
        strMessageText = err.description
    else
        strMessageText = "GRA Con-note has been added."
    end if

    Call CloseDataBase()
end function

'----------------------------------------------------------------------------------------
' Get CONNOTE from yma_gra_status
'----------------------------------------------------------------------------------------
Function getGraConnote(intGraNo)
    dim strSQL

    set rs = Server.CreateObject("ADODB.recordset")

    rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic

    Call OpenDataBase()

    strSQL = "SELECT * FROM yma_gra_status WHERE gra_no = '" & intGraNo & "'"

    'Response.Write strSQL & "<br>"

    rs.Open strSQL, conn

    if not DB_RecSetIsEmpty(rs) Then
        session("gra_connote")          = trim(rs("gra_connote"))
        session("gra_date_created")     = trim(rs("gra_date_created"))
        session("gra_created_by")       = trim(rs("gra_created_by"))
        session("gra_date_modified")    = trim(rs("gra_date_modified"))
        session("gra_modified_by")      = trim(rs("gra_modified_by"))
    end if

    call CloseDataBase()
end function

'-----------------------------------------------
' UPDATE CONNOTE
'-----------------------------------------------
function updateGraConnote(intGraNo,strConnote)
    dim strSQL

    Call OpenDataBase()

    strSQL = "UPDATE yma_gra_status SET "
    strSQL = strSQL & " gra_connote = '" & Server.HTMLEncode(strConnote) & "', "
    strSQL = strSQL & " gra_date_modified = getdate(), "
    strSQL = strSQL & " gra_modified_by = '" & session("UsrUserName") & "' "
    strSQL = strSQL & " WHERE gra_no = " & intGraNo

    on error resume next
    conn.Execute strSQL

    'response.write strSQL

    if err <> 0 then
        strMessageText = err.description
    else
        strMessageText = "GRA Con-note has been updated."
    end if

    Call CloseDataBase()
end function
%>