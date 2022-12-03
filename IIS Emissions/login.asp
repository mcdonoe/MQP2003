<!--
login.asp

Overview: This file prints an HTML form in which the user enters his/her login
      and password.  The VBScript verifies the password and displays the appropriate
      menu for that user.  If the login fails the user is prompted to try again.
Author(s): Jared McCaffree (jared@wpi.edu)
-->
<SCRIPT Language = "VBSCRIPT">

'loginButton_onClick()
'Authpr(s):	   Matthew M. Barrett
'Overview:	   Checks to see that the user has entered his name and password.
'		   Also checks to see if there is a space (" ") present in the 
'		   username and prints an error if there is
'Precondition:     The user has clicked the submit button on the login page
'Postcondition:    An error is returned if there is no username or password, or
'		   if there is a space (" ") present in the username. Otherwise,
'		   the form is submitted to the login_handler.asp page

   Sub loginButton_onClick
     Dim sErrMsg
     Dim bErrTrapped

     
     if loginForm.loginName.value = "" then
       sErrMsg = "Please enter your username"
       bErrTrapped = true
     end if

     if loginForm.frmPassword.value = "" then
       sErrMsg = "Please enter your password"
       bErrTrapped = true

     end if

     if instr(loginForm.loginName.value, " ") then
       sErrMsg = "Please enter a value username"
       bErrTrapped = true
     end if

     if bErrTrapped then
       msgBox sErrMsg
     else    
       loginForm.submit
     end if
   End Sub

</SCRIPT> 
            



<HTML>
<HEAD><TITLE>Emission's Database Login</TITLE>
<link href="../main.css" rel="stylesheet" type="text/css">
</HEAD>
<BODY>
<table  width="100%">
	<tr>
	 <td><H3>Login</H3></td>
	 <td align="right"><img src="../images/covantalogo.gif" ></td>
	</tr>
</table>
<HR>
<P>Please enter your login and password in the fields below.</P>
<FORM method=POST id=loginForm name=loginForm action="login_handler.asp">
  <TABLE>
    <TR> 
      <TD>Login Name:</TD>
      <TD><INPUT  type="text" size=20 name="loginName"></TD>
    </TR>
    <TR> 
      <TD>Password:</TD>
      <TD><INPUT size=20 type=password name="frmPassword"></TD>
    </TR>
  </TABLE>
  <INPUT type=button value="Log In" name="loginButton">
</FORM>
<HR>
<CENTER>
<BR>
<img src="../images/wpimonogram.gif" alt="A WPI Major Qualifying Project">
</CENTER>
</BODY>
</HTML>
