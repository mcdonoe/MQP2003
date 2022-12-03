<!--
removeuser.asp

Overview: This scripts generates a list of all users of the system and
          associates each user with a checkbox.  An adminstrator can remove
		  a user from the system by checking the appropriate checkbox and
		  submitting the form.
Author(s): Matthew M. Barrett
-->

<!--#include virtual="emissions/global.asi"-->

<%
   verifyUser

   Dim allUsers
   allUsers = printUsers("-c")
%>

<HTML>
<HEAD>
<TITLE>Remove User</TITLE>

<SCRIPT LANGUAGE = "VBSCRIPT">

' removeUserButton_onclick()
' Author(s):      Matthew M. Barrett
' Overview:       This subroutine is called when the Administrator presses the submit
'                   button of the Remove User form.  It checks all of the check boxes
'                   to see if any boxes or all boxes have been check. If no boxes were
'                   checked then it generates a message to the user alerting that no users
'                   were selected for removal. 
'                   the Accessible Plants list.
' Preconditions:  The Administrator clicks the submit form button after filling out the
'                   Remove User form.
' Postconditions: If the data passes inspection the form is submitted, if not an error is
'                   displayed.

Sub removeUserButton_onclick()
      for each formEntry in removeUser
         if formEntry.type = "checkbox" then
            if formEntry.Checked then
               oneChecked=true
            end if
         end if
      next
      if not oneChecked then
         errMsg = "Please select at least one user to remove."
         errTrapped = true
      end if
	  if errTrapped or not oneChecked then
          MsgBox errMsg
      else
         removeUser.submit
      end if
   End Sub

</SCRIPT>


<link href="../main.css" rel="stylesheet" type="text/css">
</HEAD>
<BODY>

   <% printHeader "Remove User" %>

<FORM action="removeuser_handler.asp" method=post id=removeUser name=removeUser
<P>&nbsp;</P>
<P>
<TABLE cellSpacing=1 cellPadding=1 border=0>
   <TR>
     <TD><% Response.Write(allUsers) %></TD></TR>
     </TABLE>
	<BR></P>
</FORM>
<INPUT type="submit" value="Delete Users" name="removeUserButton">
<BR>
   <% printFooter %>


</BODY>
</HTML>
