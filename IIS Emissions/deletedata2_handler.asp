<!--
deletedata2_handler.asp

Overview: This file actually does the deletion of the records
          from the emissions table.
Author(s): Jared F. McCaffree
           Ronald Cormier
-->

<!--#include virtual="emissions/global.asi"-->
<HTML>
<HEAD>
<%
   verifyUser

   Dim oConn          ' database connection object
   Dim deleteID
   Dim deleteDate
   deleteID = Request.Form("plantID")
   deleteDate = Request.Form("imports")

   connectDB oConn

   'get the list of imports for this particular ID
   sQueryString = "DELETE FROM emission WHERE plantID = '" & deleteID _
                  & "' AND importDate = '" & deleteDate & "';"
   oConn.Execute(sQueryString)
   'Response.Write(sQueryString)

   disconnectDB oConn
   
%>

<TITLE>Delete Data</TITLE>
<LINK href="../main.css" rel="stylesheet" type="text/css">

<BODY>

<H3>Delete Data: Complete</H3><HR>
<B>Import Deleted from system.<BR>
</B><BR>
<a href="./">Back to Menu</a>
<BR>
<BR>




</BODY>
</HTML>