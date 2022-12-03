<!--
login_handler.asp

Overview: This script receives the form information submitted by the login.asp
          form.  The script then checks to make sure the login is correct and
          sets two session variables: UserID and GroupID
Author(s): Ronald J. Cormier (rcormier@wpi.edu)
-->

<!--#include virtual="emissions/global.asi"-->


<%
   Response.Buffer = True

   Dim oConn			'database connection object
   Dim objRS			'record set object
   Dim sLoginName		'gets the users login name from the Request object
   Dim sPassword		'gets the user's password from the Request object
   Dim sQueryString		'SQL string to be executed
   Dim sTitle
   Dim sErrMsg			'Error message to be displayed
   Dim boolError		'Boolean error variable
   boolError=false
   Dim decPW			'Gets the decrypted password from the RC4.inc file

   sLoginName = Request.Form("loginName")
   sPassword  = Request.Form("frmPassword")
   sQueryString = "SELECT * FROM usertable WHERE loginID = '" & sLoginName & "'"

   connectDB oConn
   Set objRS = oConn.Execute(sQueryString)

   'check to see if user name exists. if so, then decrypt the password using the
   'EnDeCrypt function and verify the pasword the user entered is correct.

   if objRS.EOF then
      boolError = true
      sTitle    = "Error Logging In"
      sErrMsg   = "User name incorrect."
   else
      decPW = EnDeCrypt(objRS.fields("password"), objRS.fields("salt"))
      'if (objRS.fields("password") <> sPassword) then
      if (decPW <> sPassword) then
       Session("validUser") = false
       boolError = true
       sTitle    = "Error Loggin In"
       sErrMsg   = "Password incorrect."
      else
         Session("validUser") = true
         Session("UserID") = trim(objRS.fields("UserID"))
         Session("GroupID") = trim(objRS.fields("userGroup"))
         Response.Redirect("./")
      end if
   end if

   disconnectDB oConn
   set oConn = Nothing
%>

<HTML>
<HEAD><TITLE>
<%
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
   if boolError then
      Response.Write("<H3>Error: " & sErrMsg & "</H3><HR><BR>" & _
                     "<A href=./login.asp>Back</A>")
   end if
%>
</BODY>
</HTML>
