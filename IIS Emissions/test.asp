<!--
test.asp

Overview: test file
Author(s): rcormier
-->

<!--#include virtual="emissions/global.asi"-->

<HTML>
<HEAD>
<TITLE>Testing</TITLE>
</HEAD>
<BODY>

<%
   verifyUser
   
   Dim sQuery
   Dim oConn
   
   sQuery = "DELETE FROM emission"
   
   connectDB oConn
   
   oConn.Execute(sQuery)
   
   disconnectDB oConn
   
   Response.Write("All records deleted<BR>")
   Response.Write("<a href='./'>Menu</a>")
%>


</BODY>
</HTML>
