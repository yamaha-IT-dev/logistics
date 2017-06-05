<%
'----------------------------------------------------------------------------------------
' GET GRA LINE NOs: view_gra.asp
'----------------------------------------------------------------------------------------
function getGRALineNo(intGraNo)
    dim strSQL
    dim rs
	dim intLineNo
	
    call OpenBaseDataBase()
    
	strSQL = "SELECT BUHYGY FROM BFUEP "
	strSQL = strSQL & "	WHERE BUHYNO = '" & intGraNo & "'"
	strSQL = strSQL & "	ORDER BY BUHYGY"
		
	set rs = server.CreateObject("ADODB.Recordset")
    set rs = conn.execute(strSQL)
    
    if not DB_RecSetIsEmpty(rs) Then
        do until rs.EOF
        	intLineNo= trim(rs("BUHYGY"))			
			strGRALineNoList = strGRALineNoList & "<option value=" & intLineNo & ">Line " & intLineNo & "</option>"
        rs.Movenext
        loop
    end if
    
    call CloseBaseDataBase()
end function
%>