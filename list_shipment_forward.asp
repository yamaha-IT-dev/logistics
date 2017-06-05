<%
' Encrypt the UserId, Username and RoleId and redirect to the new Logistics Portal Shipment page
Dim conKey
Dim conIV

conKey = "newlogport"
conIV = Day(Date) & Month(Date) & Year(Date)
conIVnew = Year(Date) & Right("0" & Month(Date), 2) & Right("0" & Day(Date), 2)

' Create instance
Set objEncrypter = Server.CreateObject("Hyeongryeol.StringEncrypter")

objEncrypter.Key = conKey
objEncrypter.InitialVector = conIV

' Encrypt string
strEncrypted = objEncrypter.Encrypt(conIVnew & "," & Session("UsrUserID") & "," & Session("UsrUserName") & "," & Session("UsrLoginRole"))

' Encode the query string
strEncrypted = Server.URLEncode(strEncrypted)

'HACK Victor development
'If Session("UsrUserName") = "victors" Then
'    Response.Redirect "http://localhost:61217/Shipments?token=" & strEncrypted
'End If

'HACK Gandi development
'If Session("UsrUserName") = "gandig" Then
'    Response.Redirect "http://localhost:61217/Shipments?token=" & strEncrypted
'End If

' Redirect external users to the new Logistics Portal with security token
If Session("UsrLoginRole") = 3 OR Session("UsrLoginRole") = 16 OR Session("UsrLoginRole") = 17 Then
    'Response.Redirect "http://203.221.101.248:82/Shipments?token=" & strEncrypted
	Response.Redirect "http://49.255.181.168:82/Shipments?token=" & strEncrypted
	
End If

' Redirect local users to the new Logistics Portal with security token
Response.Redirect "http://172.29.64.26:82/Shipments?token=" & strEncrypted
%>