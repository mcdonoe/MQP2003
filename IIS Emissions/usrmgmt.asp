<!--
edituser.asp

Overview: This script shows all the users of the system
Author(s): Matthew M. Barrett
-->

<!--#include virtual="emissions/global.asi"-->

<%
   verifyUser

Dim allUsers      'list of all users
allUsers = printUsers("-c")
%>

<HTML>
<HEAD>
<TITLE>User Management</TITLE>

<SCRIPT LANGUAGE = "VBSCRIPT">

' editUserButton_onclick()
' Author(s):      Matthew M. Barrett
'                 Ron Cormier
' Overview:       This subroutine is called when the Administrator presses the
'                   submit button of the Edit User form.  It checks all of the
'                   check boxes to see if any boxes or all boxes have been check.
'                   If no boxes were checked then it generates a message to the
'                   user alerting that no users were selected for removal. Additionally,
'		    an error is presented if the user tries to remove the administrator
'                    account as that account can not be deleted.
' Preconditions:  The Administrator clicks the submit form button after filling out the
'                   Edit User form.
' Postconditions: If the data passes inspection the form is submitted, if not an error is
'                   displayed.
   Sub removeUserButton_onclick()
      for each formEntry in editUser
         if formEntry.type = "checkbox" then
            if formEntry.Checked then
               oneChecked=true
               if formEntry.value = 17 then
         	  errMsg = "Error: You may not remove the Administrator account"
                  errTrapped = true
               end if
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

   <% printHeader "User Management" %>

<B>Click on a name to edit that user. </B><br>
<B>To remove users, click the check box and click the Remove button</B>
<br>
<table align = right><tr><td><a href = "adduser.asp">Add New User</td></tr></table>
<BR>
<FORM action="removeuser_handler.asp" method=post id=editUser name=editUser
<P>
<TABLE cellSpacing=1 cellPadding=1 border=0>
   <TR>
     <TD><%=allUsers%></TD></TR>
     </TABLE>
	<BR></P>
</FORM>
<INPUT type="submit" value="Remove Selected User(s)" name="removeUserButton">
<BR>

   <% printFooter %>


</BODY>
</HTML>
