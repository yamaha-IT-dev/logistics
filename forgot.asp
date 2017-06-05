<!--#INCLUDE FILE = "include/connection_it.asp" -->
<%
sub checkUsername
	dim strSQL
	dim strEmail
			
	strEmail  = Trim(Request("txtEmail"))
		
	OpenDataBase()
	
	strSQL = "SELECT * FROM tbl_users WHERE email = '" & strEmail & "' "
	'response.Write strSQL	
	
	set rs = Server.CreateObject("ADODB.recordset")	
	
	rs.Open strSQL, conn
	
	if rs.EOF then
    	strMessageText = "That email was not found in our system. Please retry with the email you registered with."
    else
		dim strUsername
    	dim strPassword
		
		strUsername		= rs("username")
    	strPassword 	= rs("password")
		
		Set oMail = Server.CreateObject("CDO.Message")
		Set iConf = Server.CreateObject("CDO.Configuration")
		Set Flds = iConf.Fields
					
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.sendgrid.net"
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 10
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 'basic clear text
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "yamahamusicau"
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "str0ppy@16"
	iConf.Fields.Update
				
		emailFrom 	  	= "automailer@music.yamaha.com"				
		emailTo 	  	= strEmail
		emailSubject  	= "Logistics Portal Login"
		
		emailBodyText =	 "G'day!" & vbCrLf _
						& "" & vbCrLf _
						& "The login details as requested" & vbCrLf _ 
						& "" & vbCrLf _
						& "U: " & strUsername & vbCrLf _
						& "P: " & strPassword & vbCrLf _
						& "" & vbCrLf _
						& "This is an automated email. Please do not reply to this email."
				
		Set oMail.Configuration = iConf
		oMail.To 		= emailTo
		oMail.Cc		= emailCc
		oMail.Bcc		= emailBcc
		oMail.From 		= emailFrom
		oMail.Subject 	= emailSubject
		oMail.TextBody 	= emailBodyText
		oMail.Send
				
		Set iConf = Nothing
		Set Flds = Nothing
									
		strMessageText = "The login details have been sent to your email. Please check your inbox."
	end if	
	
	call CloseDataBase()
end sub

sub main
	if Request.ServerVariables("REQUEST_METHOD") = "POST" then
		if Trim(Request("Action")) = "Add" then
			call checkUsername
		end if
	end if
end sub

dim strMessageText

call main
%>
<!doctype html>
<html>
<head>
<link rel="stylesheet" href="css/style.css">
<link rel="stylesheet" href="bootstrap/css/bootstrap.css">
<script src="//code.jquery.com/jquery.js"></script>
<script src="bootstrap/js/bootstrap.js"></script>
<script language="JavaScript" type="text/javascript">
function trim(s)
{
  return s.replace(/^\s+|\s+$/, '');
}

function validateEmail(fld) {
    var error="";
    var tfld = trim(fld.value);                        // value of field with whitespace trimmed off
    var emailFilter = /^[^@]+@[^@.]+\.[^@]*\w\w$/ ;
    var illegalChars= /[\(\)\<\>\,\;\:\\\"\[\]]/ ;
   
    if (fld.value == "") {
        fld.style.background = 'Yellow';
        error = "Email address has not been filled in.\n";
    } else if (!emailFilter.test(tfld)) {              //test email for illegal characters
        fld.style.background = 'Yellow';
        error = "Please enter a valid email address.\n";
    } else if (fld.value.match(illegalChars)) {
        fld.style.background = 'Yellow';
        error = "Email address contains illegal characters.\n";
    } else {
        fld.style.background = 'White';
    }
    return error;
}

function validateFormOnSubmit(theForm) {
	var reason = "";
	var blnSubmit = true;

  	reason += validateEmail(theForm.txtEmail);
  
  	if (reason != "") {
    	alert(reason);    	
		blnSubmit = false;
		return false;
  	}

	if (blnSubmit == true){
        theForm.Action.value = 'Add';
		
		return true;
    }
}
</script>
<title>Logistics Portal - Forgot your login</title>
</head>
<body>
<div align="center">
  <div class="login_page">
    <h1>Forgot your login? That's cool.</h1>
    <form action="" method="post" name="form_forgot_password" id="form_forgot_password" onsubmit="return validateFormOnSubmit(this)">
      <div class="form-group">
        <label for="username">Your email:</label>
        <input type="text" class="form-control" id="txtEmail" name="txtEmail" placeholder="Email" maxlength="50" size="50">
      </div>
      <input type="hidden" name="Action" />
      <input type="submit" name="submit" id="submit" value="Retrieve" />
    </form>
    <h3><a href="./">< Back to the login</a></h3>
    <h4 style="color:green"><%= strMessageText %></h4>
  </div>
</div>
</body>
</html>