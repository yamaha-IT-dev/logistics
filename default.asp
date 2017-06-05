<!--#INCLUDE FILE = "include/connection_it.asp" -->
<!--#INCLUDE FILE = "include/AntiFixation.asp" -->
<!--#INCLUDE FILE = "class/clsLogin.asp" -->
<%
session.lcid = 2057

Sub Main()
    if(request("logout")="y")then
        Session.Abandon
        Response.Redirect("./")
    end if

    call SetSessionVariables(False)

    if Request.ServerVariables("REQUEST_METHOD") = "POST" then
        if Trim(Request("action")) = "login" then
            call SetSessionVariables(True)
            if testUserLogin then
            end if
        end if
    end if
end sub

call Main

dim strMessageText
%>
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<meta http-equiv="X-UA-Compatible" content="IE=edge">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<!--[if lt IE 9]>
  <script src="../js/html5shiv.js"></script>
  <script src="../js/respond.js"></script>
<![endif]-->
<title>Yamaha Logistics Portal</title>
<link rel="stylesheet" href="css/bootstrap.min.css">
<link rel="stylesheet" href="css/login.css">
<script src="//code.jquery.com/jquery.js"></script>
<script src="bootstrap/js/bootstrap.js"></script>
<script>
function setFocus() {
    document.forms[0].txtUsername.focus();
}

function validateForm() {
    var strUsername = document.forms[0].txtUsername.value;
    var strPassword = document.forms[0].txtPassword.value;
    var blnSubmit = true;

    if (strUsername == '') {
        alert('Please enter a username to login');
        document.forms[0].txtUsername.focus();
        blnSubmit = false;
    }

    if ((strPassword == '') && (blnSubmit == true)) {
        alert('Please enter a password to login');
        document.forms[0].txtPassword.focus();
        blnSubmit = false;
    }

    if (blnSubmit == true) {
        document.forms[0].action.value = 'login';
        return true;
    } else {
        return false;
    }
}
</script>
</head>
<body>
<div class="container">
    <form class="form-signin" role="form" name="login_form" id="login_form" method="post" action="" onsubmit="return validateForm()">
        <div class="opak">
            <h2 class="form-signin-heading">Logistics Portal</h2>
            <div class="form-group">
                <input type="text" class="form-control" id="txtUsername" name="txtUsername" maxlength="20" placeholder="Username" required autofocus>
            </div>
            <div class="form-group">
                <input type="password" class="form-control" id="txtPassword" name="txtPassword" maxlength="30" placeholder="Password" required>
            </div>
            <input type="hidden" name="action" />
            <button class="btn btn-lg btn-primary btn-block" type="submit">Sign in</button>
            <hr>
            <h4><a href="forgot/">Forgot password?</a></h4>
            <font color="green"><%= strMessageText %></font>
        </div>
    </form>
</div>
</body>
</html>