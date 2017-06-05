<%
   ' This routine is intended to provide a degree of protection
   ' against Session Fixation attacks in classic ASP

   ' Session fixation attacks are a problem in ASP, since ASP does not
   ' allow you any access to the ASPSESSIONIDxxx cookie. Even invalidating
   ' the session does not alter the value of this cookie, preventing
   ' implementation of best practice recommendations, such as
   ' issuing new session cookies when the session is authenticated, or 
   ' invalidated.

   ' The basic premise of this routine is that we create a cookie that 
   ' we CAN control, e.g. ASPFIXATION, and assign a random value to this
   ' cookie when the session is authenticated. On subsequent pages, we 
   ' check the value of this cookie against the same variable stored in
   ' the user's session. If they do not match, access is denied.
   ' When the user logs out, the session should be invalidated, and so 
   ' by default, the cookie no longer matches the value in the session.

   Private Function RandomString(l)
       Dim value, i, r
       Randomize
       For i = 0 To l
           r = Int(Rnd * 62)
           If r<10 Then
               r = r + 48
           ElseIf r<36 Then
               r = (r - 10) + 65
           Else
               r = (r - 10 - 26) + 97
           End If
           value = value & Chr(r)
       Next
       RandomString = value
   End Function

   ' This routine should be called after the user has been authenticated.
   ' It is expected that the session has been invalidated prior to this call.
   Public Sub AntiFixationInit()
       Dim value
       value = RandomString(10)
       Response.Cookies("ASPFIXATION") = value
       Session("ASPFIXATION") = value
   End Sub
   
   Public Sub AntiFixationVerify(LoginPage)
       Dim cookie_value, session_value
       cookie_value = Request.Cookies("ASPFIXATION")
       session_value = Session("ASPFIXATION")
       If cookie_value <> session_value Then
            session.abandon()
           Response.redirect(LoginPage)
       End If
   End Sub
%>