<%@ Language=VBScript %>
<!--#include file="include/connection_it.asp " -->
<%
dim rs
dim sql
dim i

dim intID
intID 	= Trim(Request("id"))

Call OpenDataBase()

set rs=server.createobject("ADODB.recordset")

sql = "SELECT * FROM yma_transfer WHERE id = '" & intID & "'"

rs.open sql,conn,1,3

Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=transfer-product-"& intID & "_list.xls"

if rs.eof <> true then
	Response.Write "<table border=1>"
	Response.Write "<tr>"
	Response.Write "<td><strong>No</strong></td>"
	Response.Write "<td><strong>Product</strong></td>"
	Response.Write "<td><strong>Qty</strong></td>"
	Response.Write "<td><strong>No of Pallets</strong></td>"
	Response.Write "<td><strong>Info</strong></td>"	
	Response.Write "<td><strong>Qty Received</strong></td>"	
	Response.Write "</tr>"   
	
	Response.Write "<tr>"
	Response.Write "<td>1</td>"
	Response.Write "<td>" & rs.fields("product_1") & "</td>"
	Response.Write "<td>" & rs.fields("qty_1") & "</td>"
	Response.Write "<td>" & rs.fields("pallet_1") & " "		
	Response.Write "<td>" & rs.fields("info_1") & "</td>"
	Response.Write "<td>" & rs.fields("received_1") & "</td>"
	Response.Write "</tr>"
	
	Response.Write "<tr>"
	Response.Write "<td>2</td>"
	Response.Write "<td>" & rs.fields("product_2") & "</td>"
	Response.Write "<td>" & rs.fields("qty_2") & "</td>"
	Response.Write "<td>" & rs.fields("pallet_2") & " "		
	Response.Write "<td>" & rs.fields("info_2") & "</td>"
	Response.Write "<td>" & rs.fields("received_2") & "</td>"
	Response.Write "</tr>"
	
	Response.Write "<tr>"
	Response.Write "<td>3</td>"
	Response.Write "<td>" & rs.fields("product_3") & "</td>"
	Response.Write "<td>" & rs.fields("qty_3") & "</td>"
	Response.Write "<td>" & rs.fields("pallet_3") & " "		
	Response.Write "<td>" & rs.fields("info_3") & "</td>"
	Response.Write "<td>" & rs.fields("received_3") & "</td>"
	Response.Write "</tr>"
	
	Response.Write "<tr>"
	Response.Write "<td>4</td>"
	Response.Write "<td>" & rs.fields("product_4") & "</td>"
	Response.Write "<td>" & rs.fields("qty_4") & "</td>"
	Response.Write "<td>" & rs.fields("pallet_4") & " "		
	Response.Write "<td>" & rs.fields("info_4") & "</td>"
	Response.Write "<td>" & rs.fields("received_4") & "</td>"
	Response.Write "</tr>"
	
	Response.Write "<tr>"
	Response.Write "<td>5</td>"
	Response.Write "<td>" & rs.fields("product_5") & "</td>"
	Response.Write "<td>" & rs.fields("qty_5") & "</td>"
	Response.Write "<td>" & rs.fields("pallet_5") & " "		
	Response.Write "<td>" & rs.fields("info_5") & "</td>"
	Response.Write "<td>" & rs.fields("received_5") & "</td>"
	Response.Write "</tr>"
	
	Response.Write "<tr>"
	Response.Write "<td>6</td>"
	Response.Write "<td>" & rs.fields("product_6") & "</td>"
	Response.Write "<td>" & rs.fields("qty_6") & "</td>"
	Response.Write "<td>" & rs.fields("pallet_6") & " "		
	Response.Write "<td>" & rs.fields("info_6") & "</td>"
	Response.Write "<td>" & rs.fields("received_6") & "</td>"
	Response.Write "</tr>"
	
	Response.Write "<tr>"
	Response.Write "<td>7</td>"
	Response.Write "<td>" & rs.fields("product_7") & "</td>"
	Response.Write "<td>" & rs.fields("qty_7") & "</td>"
	Response.Write "<td>" & rs.fields("pallet_7") & " "		
	Response.Write "<td>" & rs.fields("info_7") & "</td>"
	Response.Write "<td>" & rs.fields("received_7") & "</td>"
	Response.Write "</tr>"
	
	Response.Write "<tr>"
	Response.Write "<td>8</td>"
	Response.Write "<td>" & rs.fields("product_8") & "</td>"
	Response.Write "<td>" & rs.fields("qty_8") & "</td>"
	Response.Write "<td>" & rs.fields("pallet_8") & " "		
	Response.Write "<td>" & rs.fields("info_8") & "</td>"
	Response.Write "<td>" & rs.fields("received_8") & "</td>"
	Response.Write "</tr>"
	
	Response.Write "<tr>"
	Response.Write "<td>9</td>"
	Response.Write "<td>" & rs.fields("product_9") & "</td>"
	Response.Write "<td>" & rs.fields("qty_9") & "</td>"
	Response.Write "<td>" & rs.fields("pallet_9") & " "		
	Response.Write "<td>" & rs.fields("info_9") & "</td>"
	Response.Write "<td>" & rs.fields("received_9") & "</td>"
	Response.Write "</tr>"
	
	Response.Write "<tr>"
	Response.Write "<td>10</td>"
	Response.Write "<td>" & rs.fields("product_10") & "</td>"
	Response.Write "<td>" & rs.fields("qty_10") & "</td>"
	Response.Write "<td>" & rs.fields("pallet_10") & " "		
	Response.Write "<td>" & rs.fields("info_10") & "</td>"
	Response.Write "<td>" & rs.fields("received_10") & "</td>"
	Response.Write "</tr>"
	
	Response.Write "<tr>"
	Response.Write "<td>11</td>"
	Response.Write "<td>" & rs.fields("product_11") & "</td>"
	Response.Write "<td>" & rs.fields("qty_11") & "</td>"
	Response.Write "<td>" & rs.fields("pallet_11") & " "		
	Response.Write "<td>" & rs.fields("info_11") & "</td>"
	Response.Write "<td>" & rs.fields("received_11") & "</td>"
	Response.Write "</tr>"
	
	Response.Write "<tr>"
	Response.Write "<td>12</td>"
	Response.Write "<td>" & rs.fields("product_12") & "</td>"
	Response.Write "<td>" & rs.fields("qty_12") & "</td>"
	Response.Write "<td>" & rs.fields("pallet_12") & " "		
	Response.Write "<td>" & rs.fields("info_12") & "</td>"
	Response.Write "<td>" & rs.fields("received_12") & "</td>"
	Response.Write "</tr>"
	
	Response.Write "<tr>"
	Response.Write "<td>13</td>"
	Response.Write "<td>" & rs.fields("product_13") & "</td>"
	Response.Write "<td>" & rs.fields("qty_13") & "</td>"
	Response.Write "<td>" & rs.fields("pallet_13") & " "		
	Response.Write "<td>" & rs.fields("info_13") & "</td>"
	Response.Write "<td>" & rs.fields("received_13") & "</td>"
	Response.Write "</tr>"
	
	Response.Write "<tr>"
	Response.Write "<td>14</td>"
	Response.Write "<td>" & rs.fields("product_14") & "</td>"
	Response.Write "<td>" & rs.fields("qty_14") & "</td>"
	Response.Write "<td>" & rs.fields("pallet_14") & " "		
	Response.Write "<td>" & rs.fields("info_14") & "</td>"
	Response.Write "<td>" & rs.fields("received_14") & "</td>"
	Response.Write "</tr>"
	
	Response.Write "<tr>"
	Response.Write "<td>15</td>"
	Response.Write "<td>" & rs.fields("product_15") & "</td>"
	Response.Write "<td>" & rs.fields("qty_15") & "</td>"
	Response.Write "<td>" & rs.fields("pallet_15") & " "		
	Response.Write "<td>" & rs.fields("info_15") & "</td>"
	Response.Write "<td>" & rs.fields("received_15") & "</td>"
	Response.Write "</tr>"
	
	Response.Write "<tr>"
	Response.Write "<td>16</td>"
	Response.Write "<td>" & rs.fields("product_16") & "</td>"
	Response.Write "<td>" & rs.fields("qty_16") & "</td>"
	Response.Write "<td>" & rs.fields("pallet_16") & " "		
	Response.Write "<td>" & rs.fields("info_16") & "</td>"
	Response.Write "<td>" & rs.fields("received_16") & "</td>"
	Response.Write "</tr>"
	
	Response.Write "<tr>"
	Response.Write "<td>17</td>"
	Response.Write "<td>" & rs.fields("product_17") & "</td>"
	Response.Write "<td>" & rs.fields("qty_17") & "</td>"
	Response.Write "<td>" & rs.fields("pallet_17") & " "		
	Response.Write "<td>" & rs.fields("info_17") & "</td>"
	Response.Write "<td>" & rs.fields("received_17") & "</td>"
	Response.Write "</tr>"
	
	Response.Write "<tr>"
	Response.Write "<td>18</td>"
	Response.Write "<td>" & rs.fields("product_18") & "</td>"
	Response.Write "<td>" & rs.fields("qty_18") & "</td>"
	Response.Write "<td>" & rs.fields("pallet_18") & " "		
	Response.Write "<td>" & rs.fields("info_18") & "</td>"
	Response.Write "<td>" & rs.fields("received_18") & "</td>"
	Response.Write "</tr>"
	
	Response.Write "<tr>"
	Response.Write "<td>19</td>"
	Response.Write "<td>" & rs.fields("product_19") & "</td>"
	Response.Write "<td>" & rs.fields("qty_19") & "</td>"
	Response.Write "<td>" & rs.fields("pallet_19") & " "		
	Response.Write "<td>" & rs.fields("info_19") & "</td>"
	Response.Write "<td>" & rs.fields("received_19") & "</td>"
	Response.Write "</tr>"
	
	Response.Write "<tr>"
	Response.Write "<td>20</td>"
	Response.Write "<td>" & rs.fields("product_20") & "</td>"
	Response.Write "<td>" & rs.fields("qty_20") & "</td>"
	Response.Write "<td>" & rs.fields("pallet_20") & " "		
	Response.Write "<td>" & rs.fields("info_20") & "</td>"
	Response.Write "<td>" & rs.fields("received_20") & "</td>"
	Response.Write "</tr>"
	
	Response.Write "</table>"
end if

Call CloseDataBase()
%>