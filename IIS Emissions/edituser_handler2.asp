<!--
edituser_handler2.asp

Overview: This script receives the form information submitted by the 
          edituser_handler1.asp form.
		  The script then checks to make sure the user is unique and if so
          adds the user and his/her information to the Users table of the database.
Author(s): Jared F. McCaffree
           Ron  Cormier
-->

<!--#include virtual="emissions/global.asi"-->

<%
   verifyUser
   
   Dim oConn		' Connection Object
   Dim objRS		' Recordset object
   Dim sQueryString	' database query string
   Dim sTitle		' title of page (string)
   Dim sErrMsg		' error message to be displayed
   Dim sPlantString	' string of selected plants
   Dim b		' boolean control for plant selection
   Dim Salt		' salt key for encryption
   Dim encPW		' encrypted password returned from 'EnDeCrypt' function
   
   boolError=false
      
   ' compile the plant string
   b=false
   for each item in Request.Form
      if item = Request.Form(item) then
         if b = false then
             b=true
             sPlantString = item
         else 
             sPlantString = sPlantString & ", " & item
         end if
      end if	
   next

   connectDB oConn
   ' if there are no duplicate users add the new one
   
   Salt = GenerateSalt()
   encPW = EnDeCrypt(Request.Form("password1"), Salt)

   'Response.write("encrypted password is: " &server.urlencode(encPW))

   mySql = "UPDATE usertable SET " & _
                 "firstName='"&Request.Form("firstName")&"', " & _
			     "lastName ='"&Request.Form("lastName")&"', " & _
                                 "password = '"&encPW&"'," & _
                                 "salt = '"&Salt&"', " & _
				 "plant    ='"&sPlantString&"', " & _
				 "userGroup="&Request.Form("groupID") & _
				 " WHERE userID="&Request.Form("userID")
   '			 Response.Write mySql & "<br>"
   oConn.Execute(mySql)
   disconnectDB oConn
   
   
   	
%>

<HTML>
<HEAD><TITLE>
<%
   ' print out the proper title
   if boolError then
      Response.Write(sTitle)
   else
      Response.Write("User Updated Successfully")
   end if

%>
</TITLE>
<link href="../main.css" rel="stylesheet" type="text/css">
</HEAD>
<BODY>
<%
   ' print out an error if there was one
   if boolError then
      printHeader "Error: " & sErrMsg & "<BR> <A href=./edituser.asp>Back</A>"
   ' otherwise print out the data added to the usertable
   else
      printHeader "User Edited Successfully."
      Response.Write("<B>First Name: </B>" & Request.Form("firstName") & "<BR>")
      Response.Write("<B>Last Name: </B>" & Request.Form("lastName") & "<BR>")
      Response.Write("<B>Group: </B>" & Request.Form("groupID") & "<BR>")
      Response.Write("<B>Accessible Plants: </B>" & sPlantString & "<BR><BR>")
   end if

%>
</BODY>
</HTML>