<!--#include file="include/connection_it.asp " -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Delete Freight</title>
</head>
<body>
<%
sub deleteFreight
	dim intID
	intID = request("id")
	dim strRef
	strRef = Trim(Request("ref"))
	
    call OpenDataBase()
	
	set rs = Server.CreateObject("ADODB.recordset")
	
	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic
			
	strSQL = "DELETE FROM yma_freight WHERE id = " & intID
	
	rs.Open strSQL, conn
	
	Set rs = nothing
	
	if err <> 0 then
		strMessageText = err.description
	else 
		Response.Redirect(Request.ServerVariables("HTTP_REFERER"))
	end if 
	
    call CloseDataBase()
end sub  
    
sub main
	call deleteFreight
end sub

call main

%>
</body>