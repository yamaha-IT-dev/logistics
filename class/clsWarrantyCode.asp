<%
'----------------------------------------------------------------------------------------
' GET WARRANTY CODE LIST
'----------------------------------------------------------------------------------------

function getWarrantyCodeList
    dim strSQL
    dim rs
	dim intWarrantyCode
	
    call OpenDataBase()
    
	strSQL = "SELECT DISTINCT gra_warranty_code FROM yma_gra_warranty_code ORDER BY gra_warranty_code"
		
	set rs = server.CreateObject("ADODB.Recordset")
    set rs = conn.execute(strSQL)
    
    strWarrantyCodeList = strWarrantyCodeList & "<option value=''>...</option>"
    
    if not DB_RecSetIsEmpty(rs) Then
        do until rs.EOF
        	intWarrantyCode	= trim(rs("gra_warranty_code"))	
			
            if trim(session("report_warranty_code")) = intWarrantyCode then
                strWarrantyCodeList = strWarrantyCodeList & "<option selected value=" & intWarrantyCode & ">" & intWarrantyCode & "</option>"
            else
                strWarrantyCodeList = strWarrantyCodeList & "<option value=" & intWarrantyCode & ">" & intWarrantyCode & "</option>"
            end if
        rs.Movenext
        loop
    end if
    
    call CloseDataBase()
end function
%>