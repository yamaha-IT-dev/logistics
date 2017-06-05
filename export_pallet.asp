<%@ Language=VBScript %>
<!--#include file="include/connection_it.asp " -->
<%
dim rs
dim strSQL

Call OpenDataBase()

set rs=server.createobject("ADODB.recordset")

strSQL = "SELECT GR.*, E2ihtn+e2ihtn*e2kzrt/100+(e2ihtn*e2skkr/100) AS LIC "
strSQL = strSQL & "FROM yma_gra_report GR "
strSQL = strSQL & "LEFT JOIN AS400.S1027CFG.YGZFLIB.EF2SP "
strSQL = strSQL & "ON item = E2SOSC AND E2NGTY = YEAR(GETDATE()) AND E2NGTM = MONTH(GETDATE()) "
strSQL = strSQL & "WHERE pallet_no = '" & Request("pallet_no") & "' "
strSQL = strSQL & "ORDER BY gra_no"

rs.open strSQL,conn,1,3

Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=pallet-" & Request("pallet_no") & "-contents_list.xls"

if rs.eof <> true then
	Response.Write "<table border=1>" & vbCRLF
	Response.Write "  <tr>" & vbCRLF
	Response.Write "    <td><strong>Created by</strong></td>" & vbCRLF
	Response.Write "    <td><strong>Date Created</strong></td>" & vbCRLF
	Response.Write "    <td><strong>GRA No</strong></td>" & vbCRLF
	Response.Write "    <td><strong>Line</strong></td>" & vbCRLF
	Response.Write "    <td><strong>Item</strong></td>" & vbCRLF
	Response.Write "    <td><strong>LIC</strong></td>" & vbCRLF
	Response.Write "    <td><strong>Serial No</strong></td>" & vbCRLF
	Response.Write "    <td><strong>Dealer Code</strong></td>" & vbCRLF
	Response.Write "    <td><strong>Repair Report</strong></td>" & vbCRLF
	Response.Write "    <td><strong>Labour</strong></td>" & vbCRLF
	Response.Write "    <td><strong>Parts</strong></td>" & vbCRLF
	Response.Write "    <td><strong>GST</strong></td>" & vbCRLF
	Response.Write "    <td><strong>Total</strong></td>" & vbCRLF
	Response.Write "    <td><strong>Received</strong></td>" & vbCRLF
	Response.Write "    <td><strong>Repaired</strong></td>" & vbCRLF
	Response.Write "    <td><strong>Destination</strong></td>" & vbCRLF
	Response.Write "    <td><strong>Report Status</strong></td>" & vbCRLF
	Response.Write "  </tr>" & vbCRLF
	while not rs.eof
		Response.Write "  <tr>" & vbCRLF
		Response.Write "    <td>" & rs.fields("created_by") & "</td>" & vbCRLF
		Response.Write "    <td>" & rs.fields("date_created") & "</td>" & vbCRLF
		Response.Write "    <td>" & rs.fields("gra_no") & "</td>" & vbCRLF
		Response.Write "    <td>" & rs.fields("line_no") & "</td>" & vbCRLF
		Response.Write "    <td>" & rs.fields("item") & "</td>" & vbCRLF
		Response.Write "    <td>" & FormatNumber(rs.fields("LIC")) & "</td>" & vbCRLF
		Response.Write "    <td>" & rs.fields("serial_no") & "</td>" & vbCRLF
		Response.Write "    <td>" & rs.fields("dealer_code") & "</td>" & vbCRLF
		Response.Write "    <td>" & rs.fields("repair_report") & "</td>" & vbCRLF
		Response.Write "    <td>" & rs.fields("labour") & "</td>" & vbCRLF
		Response.Write "    <td>" & rs.fields("parts") & "</td>" & vbCRLF
		Response.Write "    <td>" & rs.fields("gst") & "</td>" & vbCRLF
		Response.Write "    <td>" & rs.fields("total") & "</td>" & vbCRLF
		Response.Write "    <td>" & rs.fields("date_received") & "</td>" & vbCRLF
		Response.Write "    <td>" & rs.fields("date_repaired") & "</td>" & vbCRLF
		Response.Write "    <td>" & rs.fields("destination") & "</td>" & vbCRLF
		Response.Write "    <td>"
		Select Case rs.fields("status")
			case 1
				Response.Write "Open"
			case 2
				Response.Write "Waiting for parts"	
			case 3
				Response.Write "To be invoiced"
			case 4
				Response.Write "Received"
			case else
		 		Response.Write "Completed / Exported"
		end select
		Response.Write "</td>" & vbCRLF
		Response.Write "  </tr>" & vbCRLF
		rs.movenext
	wend
	Response.Write "</table>"
end if

Call CloseDataBase()
%>