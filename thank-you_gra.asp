<!--#include file="include/connection_it.asp " -->
<% strSection = "gra" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>

<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Success</title>
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
</head>
<body>
<table width="100%" cellpadding="0" cellspacing="0">
  <!-- #include file="include/header.asp" -->
  <tr>
    <td class="first_content" height="400" valign="top"><p>The GRA report has been successfully added into the system.</p>
      <p>Click here to go back to <a href="view_gra.asp?id=<%= session("gra_no") %>">the previous page</a> or <a href="list_gra.asp">GRA List</a></p></td>
  </tr>
</table>
</body>
</html>