<!--
adduser_handler.asp

Overview: This script receives the form information submitted by the adduser.asp
          form.  The script then checks to make sure the user is unique and if so
          adds the user and his/her information to the Users table of the database.
Author(s): Jared F. McCaffree
           Ron  Cormier
-->

<!--#include virtual="emissions/global.asi"-->

<%
   verifyuser
   Dim oConn		' Connection Object
   Dim objRS		' Recordset object
   Dim sLoginName	' login name string
   Dim sQueryString	' database query string
   Dim sTitle		' title of page (string)
   Dim sErrMsg		' error message to be displayed
   Dim boolError	' true=error occurred when handling adduser.asp
   Dim sPlantString	' string of selected plants
   Dim b		' boolean control for plant selection
   Dim sNewUserID    ' the new id that's generated for the new user
   Dim encPW
   Dim decPW
   Dim Salt
   
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

   ' connect to database
   connectDB oConn
   
   ' get the new userID of the new user
   sNewUserID = 0
   sQueryString = "SELECT userID FROM usertable"
   Set objRS = oConn.Execute(sQueryString)
   Do While NOT objRS.EOF
      if (objRS.Fields("userID") > sNewUserID) then
	     sNewUserID = objRS.Fields("userID")
      end if
	  objRS.MoveNext
   loop
   sNewUserID = CInt(sNewUserID + 1)
   
   ' store login name and compile query string
   sLoginName = Request.Form("userName")
   sQueryString = "SELECT * FROM usertable WHERE loginID='" & sLoginName & "'"

   ' Execute the query string and return the resulting recordset containing
   '   all users that have the login name that's being added (test for duplicates).
   Set objRS = oConn.Execute(sQueryString)
   if not objRS.EOF then	' if we don't get an empty set back
      boolError=true		'  there's an error
      sTitle = "Error Adding User"
      sErrMsg = "Must enter a user name that does not already exist."
   else
      ' if there are no duplicate users add the new one
     Salt = GenerateSalt()
     encPw = EnDeCrypt(Request.Form("password1"), Salt)
   '  response.write("The encrypted password is " & server.urlencode(encPW) & "<br>")

   '  decPW = EnDeCrypt(encPW, Salt)
   '  response.write("The decrypted password is " & decPW)
     oConn.Execute("INSERT INTO usertable(firstName,lastName,userGroup,password,salt,loginID,plant,userID) " & _
                     "Values('" & Request.Form("firstName") & "', '" & _
                     Request.Form("lastName") & "', '" & Request.Form("groupID")& _
                     "', '" & encPW & "', '" & Salt & _
                      "', '" & Request.Form("userName") & "', '" & sPlantString & _ 
                      "', '"  & sNewUserID & "')")
   end if
   disconnectDB oConn
%>

<HTML>
<HEAD><TITLE>
<%
   ' print out the proper title
   if boolError then
      Response.Write(sTitle)
   else
      Response.Write("User Added Successfully")
   end if

%>
</TITLE>
<link href="../main.css" rel="stylesheet" type="text/css">
</HEAD>
<BODY>
<%
   ' print out an error if there was one
   if boolError then
      printHeader ("<H3>Error: " & sErrMsg & "</H3><BR>" & _
                     "<A href=./adduser.asp>Back</A>")
   ' otherwise print out the data added to the usertable
   else
      printHeader "User Added Successfully"
      Response.Write("<B>User Name: </B>" & Request.Form("userName") & "<BR>")
      Response.Write("<B>First Name: </B>" & Request.Form("firstName") & "<BR>")
      Response.Write("<B>Last Name: </B>" & Request.Form("lastName") & "<BR>")
      Response.Write("<B>Group: </B>" & Request.Form("groupID") & "<BR>")
      Response.Write("<B>Accessible Plants: </B>" & sPlantString & "<BR><BR>")
   end if

%>
</BODY>
</HTML>