<!--
changepassword.asp

Overview: This file allows a user to change his or her password.

Author(s): Jared F. McCaffree, Matthew M. Barrett
-->

<!--#include virtual="emissions/global.asi"-->

<%
   verifyUser
%>

<HTML>
<TITLE>Change Your Password</TITLE>
<HEAD>

<SCRIPT Language = "VBScript">



' changePwordButton_onclick()
' Author(s):      Matthew M. Barrett
' Overview:       This subroutine is called when the user presses the submit
'                   button of the Change Password form.  It checks that the current
'                   password entered is correct, and that the new password was correctly
'                   entered two times.
' Preconditions:  The user clicks the submit form button after filling out the
'                   Change Password Form  form.
' Postconditions: If the data passes inspection the form is submitted and the password
'                   is changed. If not an error is displayed.

   Sub changePwordButton_onclick()

     Dim errTrapped		'boolean error value    
     if form1.oldpassword.value = "" then
	errMsg = "Please enter your current password"	
	errTrapped = true
     End If

     if form1.newpassword1.value = "" then
	errMsg = "Please enter your new password in the New Password box"
	errTrapped = true
     End If

     if form1.newpassword2.value = "" then
	errMsg = "Please enter you new password in the New Password Again box"
     End If

     if form1.newpassword1.value <> form1.newpassword2.value then
	errMsg = "New passwords do not match. Please enter new matching passwords"
	errTrapped = true
     End If

     if errTrapped then
	Msgbox errMsg 
     else
	form1.submit
     End IF

  End Sub

</SCRIPT>

<LINK href="../main.css" rel="stylesheet" type="text/css">
</HEAD>

<BODY>
<% printHeader "Change Your Password" %>

<B>Fill out the form to change your password</B>
<BR>

<FORM action="changepassword_handler.asp" method=post id=form1 name=form1>
<TABLE cellspacing=1 cellpadding=1 border=0>
<TABLE>
   <TR>
      <TD>Old Password:</TD>
      <TD><INPUT id="oldpassword" name="oldpassword" type="password"></TD>
   </TR><TR>
      <TD>New Password:</TD>
      <TD><INPUT id="newpassword1" name="newpassword1" type="password"></TD>
   </TR><TR>
      <TD>New Password Again:</TD>
      <TD><INPUT id="newpassword2" name="newpassword2" type="password"></TD>
   </TR>
</TABLE>
<INPUT type="button" value="Change Password" id="changePwordButton" 
  name="changePwordButton">
</FORM>

<% printFooter %>

</BODY>
</HTML>