<%
'-----------------------------------------------
' LIST COMMENTS
'-----------------------------------------------
function listComments(intID,intTypeID)
    dim strSQL
    dim intRecordCount

    call OpenDataBase()

    set rs = Server.CreateObject("ADODB.recordset")

    rs.CursorLocation = 3   'adUseClient
    rs.CursorType = 3       'adOpenStatic
    rs.PageSize = 200

    strSQL = "SELECT * FROM tbl_comments "
    strSQL = strSQL & "	WHERE associated_id = '" & intID & "' "
    strSQL = strSQL & "		AND comment_type = '" & intTypeID & "' "
    strSQL = strSQL & "	ORDER BY comment_date"

    rs.Open strSQL, conn

    intRecordCount = rs.recordcount

    strCommentsList = ""

    if not DB_RecSetIsEmpty(rs) Then

        For intRecord = 1 To rs.PageSize

            strCommentsList = strCommentsList & "<tr><td class=""comment_column"">"	& trim(rs("comments")) & "</td></tr>"
            strCommentsList = strCommentsList & "<tr><td class=""comment_content""><strong>" & trim(rs("comment_by")) & "</strong> - " & WeekDayName(WeekDay(rs("comment_date"))) & ", " & FormatDateTime(rs("comment_date"),1) & " at " & FormatDateTime(rs("comment_date"),3) & "</td></tr>"

            rs.movenext

            If rs.EOF Then Exit For
        next
    else
        strCommentsList = "<tr><td>&nbsp;</td></tr>"
    end if

    strCommentsList = strCommentsList & "<tr>"

    call CloseDataBase()
end function

'-----------------------------------------------
' ADD COMMENT
'-----------------------------------------------
function addComment(intID,intTypeID)
    dim strSQL

    dim strComment
    strComment      = Replace(Request.Form("txtComment"),"'","''")

    Call OpenDataBase()

    strSQL = "INSERT INTO tbl_comments (comment_type, comments, associated_id, comment_by) VALUES ('" & intTypeID & "',"
    strSQL = strSQL & " '" & Server.HTMLEncode(strComment) & "',"
    strSQL = strSQL & " '" & intID & "',"
    strSQL = strSQL & " '" & lcase(session("UsrUserName")) & "')"

    on error resume next
    conn.Execute strSQL

    if err <> 0 then
        strMessageText = err.description
    else
        strMessageText = "<div class=""notification_text""><img src=""images/icon_check.png""> Comment has been added.</div>"
    end if

    Call CloseDataBase()
end function
%>