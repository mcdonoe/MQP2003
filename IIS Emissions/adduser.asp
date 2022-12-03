<!--
adduser.asp

Overview: This file displays an add user form and upon submittion of the form verifies
           all entered information and submits the form to be processed by the
           adduser_handler.asp script.
Author(s): Jared F. McCaffree
           Ron Cormier
-->

<!--#include virtual="emissions/global.asi"-->

<SCRIPT Language="VBScript">

' cmdSelectAllPlants_onclick()
' Author(s):      Jared F. McCaffree
' Overview:       When the "Select All" button is clicked this subroutine is called.
'                   This subroutine sets the checked value of each checkbox on the page
'                   to true.

   Sub cmdSelectAllPlants_onclick

      dim oFormEntry                ' for each loop value

      for each oFormEntry in form1
         if oFormEntry.type = "checkbox" then
            oFormEntry.checked = true
         end if
      next
   
   End Sub
</SCRIPT>
<%
   verifyUser
%>

<HTML>
<HEAD>
<TITLE>Add User</TITLE>

<%

Dim allGroups   'holds all the groups a user can be
Dim allPlants   'holds all the plants a user can be assoc. w/
Dim iUserID
allGroups = printGroups("")
allPlants = printPlants("-c", "")

iUserID = Request("uid")
'Response.write("the user id is " & sUsrName)
isCopy = false
if iUserID <> "" then
   Dim oConn
   Dim userInfo
   Dim iUGroup
   isCopy = true
   connectDB oConn
   set userInfo = oConn.execute("SELECT * FROM usertable WHERE userID=" & iUserID)
   sPlant = userInfo.Fields("plant")
   allPlants = printPlants("-c", sPlant)
   allGroups = printGroups(userInfo.Fields("userGroup"))
end if 
%>

<SCRIPT Language="VBScript">
 

' cmdSubmitButton_onclick()
' Author(s):      Jared F. McCaffree
' Overview:       This subroutine is called when the Administrator presses the submit
'                   button of the Add User form.  It checks all of the text boxes
'                   for valid input and displays an error if the Administrator attempts
'                   to submit the form without filling in a required field.  The form also
'                   checks to make sure the Administrator checks at least one plant from
'                   the Accessible Plants list.
' Preconditions:  The Administrator clicks the submit form button after filling out the
'                   Add User form.
' Postconditions: If the data passes inspection the form is submitted, if not an error is
'                   displayed.

   Sub cmdSubmitButton_onclick

      dim oneChecked	' Boolean for checking validity of plant selection
      dim errMsg        ' String of the error message to be displayed
      dim errTrapped    ' Boolean, set true if an error occurrs

      oneChecked=false
      errTrapped=false

      if form1.userName.value = "" then
         errMsg = "Please enter a user name."
         errTrapped = true
      end if
      if Len(form1.password1.value) < 5 then
         errMsg = "Please enter a password more than 5 characters."
         errTrapped = true
      end if
      if Len(form1.password2.value) < 5 then
         errMsg = "Please make sure both passwords are more than 5 characters."
         errTrapped = true
      end if
      if form1.firstName.value = "" then
         errMsg = "Please enter the user's first name."
         errTrapped = true
      end if
      if form1.lastName.value = "" then
         errMsg = "Please enter the user's last name."
         errTrapped = true
      end if
      if form1.password1.value <> form1.password2.value then
         errMsg = "Passwords do not match, please re-enter the password again."
         errMsg = errMsg & "  Passwords must be more than 5 characters."
         errTrapped = true
         form1.password1.value=""
         form1.password2.value=""
      end if
      for each formEntry in form1
         if formEntry.type = "checkbox" then
            if formEntry.Checked then
               oneChecked=true
            end if
         end if
      next
      if not oneChecked then
         errMsg = "Please select at least one accessible plant."
         errTrapped = true
      end if
      if errTrapped or not oneChecked then
         MsgBox errMsg
      else
         form1.submit
      end if
   End Sub
</SCRIPT>

<link href="../main.css" rel="stylesheet" type="text/css">
</HEAD>
<BODY>

   <% printHeader "Add User" %>

<B>Fill in the form with the new user's information</B>
<BR>
<FORM action="adduser_handler.asp" method=post id=form1 name=form1>
<P>
<TABLE cellSpacing=1 cellPadding=1 border=0>
   <TR>
     <TD>User Name</TD>
     
     <TD><INPUT id="userName" name="userName"     
       <%if isCopy = true then
       Response.write(" value = 'Copy of " & userInfo.Fields("loginID") & "'") 
       end if %>
     ></TD></TR>
     <TD>Password</TD>
     <TD><INPUT type="password" id="password1" name="password1"></TD></TR>
   <TR>
     <TD>Password Again</TD>
     <TD><INPUT type="password" id="password2" name="password2"></TD></TR>
   <TR>
     <TD>First Name</TD>
     <TD><INPUT id="firstName" name="firstName" 
     <%if isCopy = true then
       Response.write(" value = '" & userInfo.Fields("firstName") & "'") 
       end if %>
     ></TD></TR>
   <TR>
     <TD>Last Name</TD>
     <TD><INPUT id="lastName" name="lastName"
     <%if isCopy = true then
       Response.write(" value = '" & userInfo.Fields("lastName") & "'") 
       end if %>
     ></TD></TR>
   <TR>
     <TD>Group</TD>
     <TD><%= allGroups %><TD></TR>
   <TR>
     <TD>Accessible Plants</TD>
     <TD><%= allPlants %>
     <INPUT type="button" value="Select All" id="cmdSelectAllPlants" 
               name="cmdSelectAllPlants"><BR>
     </TD></TR></TABLE>
	<BR></P>
</FORM>
<INPUT type="button" value="Add User" id="cmdSubmitButton" name="cmdSubmitButton">
<BR>

   <% printFooter %>


</BODY>
</HTML>
