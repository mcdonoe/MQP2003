<!--
edituser.asp

Overview: This script shows all the users of the system
Author(s): Matthew M. Barrett
-->

<!--#include virtual="emissions/global.asi"-->

<%
   verifyUser

Dim allUsers      'list of all users
allUsers = printUsers("-r")
%>

<HTML>
<HEAD>
<TITLE>Edit User</TITLE>

<SCRIPT LANGUAGE = "VBSCRIPT">

' editUserButton_onclick()
' Author(s):      Matthew M. Barrett
'                 Ron Cormier
' Overview:       This subroutine is called when the Administrator presses the
'                   submit button of the Edit User form.  It checks all of the
'                   check boxes to see if any boxes or all boxes have been check.
'                   If no boxes were checked then it generates a message to the
'                   user alerting that no users were selected for removal.
' Preconditions:  The Administrator clicks the submit form button after filling out the
'                   Edit User form.
' Postconditions: If the data passes inspection the form is submitted, if not an error is
'                   displayed.

Sub editUserButton_onclick()
      for each formEntry in editUser
         if formEntry.type = "radio" then
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
         editUser.submit
      end if
   End Sub

</SCRIPT>


<link href="../main.css" rel="stylesheet" type="text/css">
</HEAD>
<BODY>

   <% printHeader "Edit User" %>

<B>Select User to edit</B>
<BR>
<FORM action="edituser_handler1.asp" method=post id=editUser name="editUser">
<P>
<TABLE cellSpacing=1 cellPadding=1 border=0>
   <TR>
     <TD><%=allUsers%></TD></TR>
     </TABLE>
	<BR></P>
</FORM>
<INPUT type="submit" value="Edit User" name="editUserButton">
<BR>

   <% printFooter %>


</BODY>
</HTML>
