<!--
changepword_handler.asp

Overview: This scripts receives the form information submitted by the 
          changepword.asp form.  The script then checks to see that the 
          user has entered his or her correct current password.  If so, 
          then the user's password is changed to the value entered in 
          changepword.asp.
Author(s): Matthew M. Barrett
-->

<!--#include virtual="emissions/global.asi"-->

<%
   verifyUser
   
   Dim oConn
   Dim objRS   		' Recordset object (result of SQL statement)
   Dim boolBadCurPwd	'boolean var; set to true if current password entered 	
			'in changepword.asp doesn't match actual current pwd

   sPassword  = Request.Form("oldpassword")
   curLogin = Session("UserID")
   Dim sQueryString 	'SQL string used to query database pwd
   Dim sUpdateString	'SQL string to update database w/ new pwd
   Dim Salt 		'salt used to encrypt password
   Dim encPW		'holds encrytped pwd after calling EnDeCrypt function
   Dim sCurPW		'holds current password
   connectDB oConn
   
   sQueryString = "SELECT * FROM userTable WHERE userID ='" & curLogin &"'"

   Set objRS = oConn.Execute(sQueryString)

   'Decrypt the current password for the user
   sCurPW = EnDeCrypt(objRS.fields("password"), objRS.fields("salt"))

   'If the decrypted password matches the password the user has entered in the
   'Current Password box then a Salt key is generated, and the new password
   'is encrypted using that Salt. The Salt and encrypted password are then stored
   'in the database.

   if (sCurPW <> sPassword) then
     errMsg = "Error: The current password you entered is incorrect."
     boolCurPwdBad = true
   else
     Salt = GenerateSalt()
     encPW = EnDeCrypt(Request.Form("newpassword2"), Salt)
     sUpdateString = "UPDATE userTable set password = '" & encPW & _ 
                     "', " & "salt = '" &Salt &"' " & _
                     "WHERE userID ='"&curLogin&"'"
     Set objRS = oConn.Execute(sUpdateString)
     result = "The password was successfully changed."
   End IF
   
   disconnectDB oConn

%>


<HTML>
<HEAD>
<TITLE Change Password Result> </TITLE>

<link href="../main.css" rel="stylesheet" type="text/css">
</HEAD>
<BODY>



<%
   if boolCurPwdBad = true then
        printHeader "<H3>"& errMsg & "</H3><BR>" & _
        "<A href=./changepword.asp>Back</A>"
   else
        printHeader "<H3> The password has been successfully changed. </H3><BR>"
   end If
%>


</BODY>
</HTML>
