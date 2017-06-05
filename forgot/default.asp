<!--#include file="../../include/connection.asp " -->
<!--#include file="../class/clsPassword.asp " -->
<%
sub main
	if Request.ServerVariables("REQUEST_METHOD") = "POST" then
		if trim(request("action")) = "retrieve" then
			dim strUsername
			strUsername = Trim(Request("txtEmail"))
			
			call checkUsername(strUsername)
		end if
	end if
end sub

dim strMessageText
call main
%>
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<meta http-equiv="X-UA-Compatible" content="IE=edge">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<!--[if lt IE 9]>
  <script src="../../js/html5shiv.js"></script>
  <script src="../../js/respond.js"></script>
<![endif]-->
<title>Forgot Password</title>
<link rel="stylesheet" href="../../css/bootstrap.min.css">
<link rel="stylesheet" href="../css/login.css">
<script src="//code.jquery.com/jquery.js"></script>
<script src="../../bootstrap/js/bootstrap.js"></script>
<script src="../include/generic_form_validations.js"></script>
<script>
function validateFormOnSubmit(theForm) {
	var reason = "";
	var blnSubmit = true;

  	reason += validateEmptyField(theForm.txtEmail,"Username");
  
  	if (reason != "") {
    	alert(reason);    	
		blnSubmit = false;
		return false;
  	}

	if (blnSubmit == true){
        theForm.action.value = 'retrieve';
		
		return true;
    }
}
</script>
</head>
<body>
<div class="container">
  <form class="form-signin" role="form" action="" method="post" name="form_forgot_password" id="form_forgot_password" onsubmit="return validateFormOnSubmit(this)">
    <div class="opak">
      <h2>Forgot password?</h2>
      <div class="form-group">
        <label for="txtEmail">Please enter your email<font color="red">*</font>:</label>
        <input type="email" class="form-control" name="txtEmail" id="txtEmail" placeholder="Email" maxlength="50" size="34" />
      </div>
      <div class="form-group">
        <input type="hidden" name="action" />
        <input type="submit" name="submit" id="submit" class="btn btn-lg btn-primary btn-block" value="Submit" />
      </div>
      <hr>
      <h4><a href="../">< Back to login</a></h4>
    </div>
    <br>
    <%= strMessageText %>
  </form>
</div>
</body>
</html>