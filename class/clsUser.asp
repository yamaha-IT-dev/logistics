<%
'-----------------------------------------------
' ADD NEW USER
'-----------------------------------------------
function addUser(strUsername,strFirstname,strLastname,strEmail,strDepartment,intAdmin,intStatus)
	dim strSQL
	
	Call OpenDataBase()
		
	strSQL = "INSERT INTO tbl_users ("
	strSQL = strSQL & "username, "
	strSQL = strSQL & "firstname, "
	strSQL = strSQL & "lastname, "
	strSQL = strSQL & "email, "
	strSQL = strSQL & "department, "
	strSQL = strSQL & "admin, "
	strSQL = strSQL & "created_by "
	strSQL = strSQL & ") VALUES ("
	strSQL = strSQL & " '" & strUsername & "',"
	strSQL = strSQL & " '" & Server.HTMLEncode(strFirstname) & "',"
	strSQL = strSQL & " '" & Server.HTMLEncode(strLastname) & "',"
	strSQL = strSQL & " '" & strEmail & "',"
	strSQL = strSQL & " '" & strDepartment & "',"
	strSQL = strSQL & " '" & intAdmin & "',"
	strSQL = strSQL & " '" & session("logged_username") & "' "
	strSQL = strSQL & ")"
	
	'response.Write strSQL
	on error resume next
	conn.Execute strSQL
	
	if err <> 0 then
		strMessageText = err.description
	else
		strMessageText = "The record has been added."
	end if
	
	Call CloseDataBase()
end function

'-----------------------------------------------
' GET USER DETAILS (using username)
'-----------------------------------------------
Function getUserDetails(strUsername)
	dim rs
	dim strSQL
	
    call OpenDataBase()

	set rs = Server.CreateObject("ADODB.recordset")

	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic

	strSQL = "SELECT * FROM tbl_users "
	strSQL = strSQL & " WHERE username = '" & strUsername & "'"

	rs.Open strSQL, conn
	'Response.Write strSQL
	
    if not DB_RecSetIsEmpty(rs) Then
		session("user_id") 		= rs("user_id")
		session("username") 	= rs("username")
		session("firstname") 	= rs("firstname")
		session("lastname") 	= rs("lastname")
		session("email") 		= rs("email")
		session("department")	= rs("department")
		session("status") 		= rs("status")
		session("admin") 		= rs("admin")
    end if

    call CloseDataBase()
end function

'-----------------------------------------------
' GET USER DETAILS
'-----------------------------------------------
Function getUser(intUserID)
	dim rs
	dim strSQL
	
    call OpenDataBase()

	set rs = Server.CreateObject("ADODB.recordset")

	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic

	strSQL = "SELECT * FROM tbl_users "
	strSQL = strSQL & " WHERE user_id = '" & intUserID & "'"

	rs.Open strSQL, conn
	'Response.Write strSQL
	
    if not DB_RecSetIsEmpty(rs) Then
		session("user_id") 			= rs("user_id")
		session("user_username") 	= rs("username")
		session("user_firstname") 	= rs("firstname")
		session("user_lastname") 	= rs("lastname")
		session("user_email") 		= rs("email")
		session("user_department")	= rs("department")
		session("user_status") 		= rs("status")
		session("user_admin") 		= rs("admin")
		session("user_date_created")= rs("date_created")
		session("user_created_by") 	= rs("created_by")
		session("user_date_modified")= rs("date_modified")
		session("user_modified_by") = rs("modified_by")
    end if

    call CloseDataBase()
end function

'-----------------------------------------------
' UPDATE USER
'-----------------------------------------------
function updateUser(intUserID,strUsername,strFirstname,strLastname,strEmail,strDepartment,intAdmin,intStatus)
	dim strSQL
	
	Call OpenDataBase()
		
	strSQL = "UPDATE tbl_users SET "
	strSQL = strSQL & "username = '" & strUsername & "',"
	strSQL = strSQL & "firstname = '" & Server.HTMLEncode(strFirstname) & "',"
	strSQL = strSQL & "lastname = '" & Server.HTMLEncode(strLastname) & "',"
	strSQL = strSQL & "email = '" & strEmail & "',"
	strSQL = strSQL & "department = '" & strDepartment & "',"
	strSQL = strSQL & "admin = '" & intAdmin & "',"
	strSQL = strSQL & "status = '" & intStatus & "',"
	strSQL = strSQL & "date_modified = getdate(),"
	strSQL = strSQL & "modified_by = '" & session("logged_username") & "' "
	strSQL = strSQL & "	WHERE user_id = " & intUserID
	
	'response.Write strSQL
	on error resume next
	conn.Execute strSQL
	
	if err <> 0 then
		strMessageText = err.description
	else
		strMessageText = "The record has been updated."
	end if
	
	Call CloseDataBase()
end function

%>