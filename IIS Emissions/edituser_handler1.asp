<!--
edituser_handler1.asp

Overview: This scripts receives the form information submitted by the
            edituser.asp form, which is the list of users to be edited.  The
			script displays each of the user's information in editable form.
Author(s): Matthew M. Barrett
           Ron Cormier
-->

<!--#include virtual="emissions/global.asi"-->

<HTML>
<HEAD>

<%
   verifyUser       'verify user is logged in

   Dim strEdit  	'contains the id of user to edit
   Dim objRS   		'Recordset object (result of SQL SELECT statement)
   Dim oConn
   Dim sQueryString
   Dim boolUserFound   'was the user found in the system?
   Dim errMessage

   'the following will be used to fill in the form
   Dim sFirstName
   Dim sLastName
   Dim sGroup
   Dim sPlant
   Dim sCurPwd
   Dim sUName
   connectDB oConn
   
   strEdit = Request("uid")
   sQueryString="SELECT * FROM usertable WHERE userID = "&strEdit
   Set objRS = oConn.Execute(sQueryString)

   if (NOT objRS.EOF) then
      boolUserFound = true
	  errMessage = "None"
	  sFirstName = Trim(objRS.Fields("firstName"))
	  sLastName = Trim(objRS.Fields("lastName"))
	  sGroup = objRS.Fields("userGroup")
          sUName = objRS.Fields("loginID")
	  sPlant = objRS.Fields("plant")
	  sCurPwd = objRS.Fields("password")
	  sGroup = printGroups(sGroup)
	  sPlant = printPlants("-c", sPlant)
   else
      boolUserFound = false
	  errMessage = "Error: User not found"
   end if
   disconnectDB oConn
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


      if form1.firstName.value = "" then
         errMsg = "Please enter the user's first name."
         errTrapped = true
      end if
      if form1.lastName.value = "" then
         errMsg = "Please enter the user's last name."
         errTrapped = true
      end if
      for each formEntry in form1
         if formEntry.type = "checkbox" then
            if formEntry.Checked then
               oneChecked=true
            end if
         end if
      next
      if form1.password1.value <> form1.password2.value then
          errMsg = "The passwords you have entered do not match. Please re-enter passwords."
          errTrapped = true
      end if
      if Len(form1.password1.value) < 5 then
	  errMsg = "Please enter a password in to the 'password' and 'verify password' boxes."
          errMsg = errMsg & "  Passwords must be more than 5 characters."
          errTrapped = true
      end if
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
<%
   'if the user was found, display the form
   if (boolUserFound = true) then
%>
  <% printHeader "Edit the User's Information" %>

<A HREF='usrmgmt.asp'>Back</A>
<FORM action="adduser.asp?uid=<%=strEdit%>" method="POST" id=formc name=formc>
<INPUT TYPE = "submit" value="Copy This User" id="copy" name="copy">
</FORM>

<FORM action="edituser_handler2.asp" method=post id=form1 name=form1>
<INPUT type=hidden name=userID value=<%=strEdit%>>
<P>
<TABLE cellSpacing=1 cellPadding=1 border=0>
   <TR>
     <TD>First Name</TD>
     <TD><INPUT id="firstName" name="firstName" value="<%=sFirstName%>"></TD></TR>
   <TR>
     <TD>Last Name</TD>
     <TD><INPUT id="lastName" name="lastName" value="<%=sLastName%>"></TD></TR>
   <TR>
     <TD>User Name</TD>
     <TD><%=sUName%></TD></TR>
   <TR>
     <TD>Password</TD>
     <TD><INPUT type = "password" id="password1" name="password1" value="<%=sCurPwd%>"></TD></TR>
   <TR>
     <TD>Verify Password</TD>
     <TD><INPUT type = "password" id="password2" name="password2" value="<%=sCurPwd%>"></TD></TR>
   <TR>
     <TD>Group</TD>
     <TD><%= sGroup %><TD></TR>
   <TR>
     <TD>Accessible Plants</TD>
     <TD><%= sPlant %></TD>
   </TR>
</TABLE>
	<BR></P>
</FORM>
<INPUT type="button" value="Edit User" id="cmdSubmitButton" name="cmdSubmitButton">
<INPUT type="button" value="Cancel Edit" id="cmdCancelButton" name="cmdCancelButton"
onclick="location='./'">
<%
   else
      Response.Write(errMessage)
	  Response.Write("<BR><BR><A HREF='./'>Main Page")
   end if
%>
</BODY>
</HTML>