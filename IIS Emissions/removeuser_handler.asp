<!--
removeuser_handler.asp

Overview: This scripts receives the form information submitted by the removeuser.asp form.
	  The script then retrieves and prints a list of all users from the usertable and
	  places a checkbox next to the user's name. The user can then be removed from
	  the system by clicking the checkboxes and submitting the form.
Author(s): Matthew M. Barrett
-->

<!--#include virtual="emissions/global.asi"-->

<HTML>
<HEAD>

<%
   verifyUser
   
   Dim strDeleteList 	    		     'will contain list of users to be deleted
   Dim objRS   		                     'Recordset object (result of SQL SELECT statement)
   Dim oConn		                     'database connection object

   connectDB oConn
   
   strDeleteList=Request("userIDs")          'get checked information from previous page
   objRS="DELETE FROM usertable WHERE userID IN("&strDeleteList&")"
   oConn.Execute objRS			    'Execute the SQL delete command

   disconnectDB oConn
   
   'display verification work
   Dim numRem                               'number of users removed
   Dim indivUsers                           'array of user id nums
   indivUsers=Split(strDeleteList)
   For Each user In indivUsers		    'count how many users were deleted
      numRem = numRem + 1
   Next
   
%>

<link href="../main.css" rel="stylesheet" type="text/css">

</HEAD>

<BODY>

   <% printHeader numRem & "</B> user(s) removed from the system" %>
</BODY>
</HTML>