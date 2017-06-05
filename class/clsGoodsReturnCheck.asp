<%
function checkAV(strGRA)
	dim strSQL
	dim rs
	
	call OpenLiveDataBase()
	
	strSQL = "SELECT * FROM yma_gra WHERE gra_no = '" & strGRA & "'"
	
	'response.Write strSQL	
	
	set rs = Server.CreateObject("ADODB.recordset")
	
	rs.Open strSQL, conn
	
	if rs.EOF then    	
		session("gra_AV_check") = 0
    else		
		session("gra_AV_check") = 1
	end if
	
	call CloseDataBase()
end function

function checkMPD(strGRA)
	dim strSQL
	dim rs
	
	call OpenLiveDataBase()
	
	strSQL = "SELECT * FROM tbl_gra_mpd WHERE gra_no = '" & strGRA & "'"
	
	'response.Write strSQL	
	
	set rs = Server.CreateObject("ADODB.recordset")
	
	rs.Open strSQL, conn
	
	if rs.EOF then    	
		session("gra_MPD_check") = 0
    else		
		session("gra_MPD_check") = 1
	end if
	
	call CloseDataBase()
end function
%>