<!--
printemissions.asp

Overview: 
Author(s): Jared F. McCaffree

-->

<!--#include virtual="emissions/global.asi"-->

<HTML>
<BODY>

<%


   Dim oConn		' Connection Object
   Dim objRS		' Recordset object
   Dim sQueryString	' database query string
 
   sQueryString = "SELECT * FROM emission"

   connectDB oConn
   Set objRS = oConn.Execute(sQueryString)

   do while not objRS.EOF
      response.write(objRS.Fields("plantID") & "<br>")
'       response.write("shit")
      objRS.MoveNext
   loop

%>
</BODY>
</HTML>