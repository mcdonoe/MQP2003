<!--
deletedata_handler.asp

Overview: This file gets the plant that the user wants to delete from
          and displays the import dates for that plant.  The date chosen
          will be deleted from the emissions table.
Author(s): Jared F. McCaffree
           Ronald Cormier
-->

<!--#include virtual="emissions/global.asi"-->
<HTML>
<HEAD>
<%
   verifyUser

   Dim oConn          ' database connection object
   Dim objRS
   Dim sQueryString
   Dim sPlantID
   Dim sImportDates   ' HTML for user to select import to delete
   Dim sOneDate

   connectDB oConn

   ' get the plant id
   sQueryString = "SELECT * FROM plant WHERE plantName = '" & _
   Request.Form("plants") & "';"
   Set objRS = oConn.Execute(sQueryString)
   sPlantID = trim(objRS.Fields("ID"))

   'get the list of imports for this particular ID
   sQueryString = "SELECT importDate FROM emission WHERE plantID = "& _
   CInt(sPlantID) & " GROUP BY importDate;"
   Set objRS = oConn.Execute(sQueryString)

   sImportDates = ""
   if NOT objRS.EOF then
      do while NOT objRS.EOF
         sOneDate = trim(objRS.Fields("importDate"))
         sImportDates = sImportDates & "<INPUT type=radio name=imports" & _
                        " value='" & sOneDate & "'>" & _
                        sOneDate & "</input>" & vbcrlf & "<BR>"
         objRS.MoveNext
      loop
   else
      sImportDates = "No imports on file"
   end if

   disconnectDB oConn
   
%>

<TITLE>Delete Data</TITLE>
<LINK href="../main.css" rel="stylesheet" type="text/css">

<BODY>

<H3>Delete Data Step 2: Select Import</H3><HR>
<B>The second step of the data delete process is to select which import 
to delete from this plant.<BR>
</B><BR>
<a href="./">Back to Menu</a>
<BR>
<BR>
<FORM METHOD="POST" ACTION="deletedata2_handler.asp">
   <TABLE BORDER=0>
      <TR><TD><B>Select import:</B></TD>
      <TD><% =sImportDates %></TD></TR>
      <TR><TD colspan=2><INPUT type = submit value="Remove Imports">
      </TD></TR>
   </TABLE>
   <input type=hidden name='plantID' value='<%=sPlantID%>'>
</FORM>


</BODY>
</HTML>