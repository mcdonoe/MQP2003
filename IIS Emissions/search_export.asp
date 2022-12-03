<!--
search_export.asp

Overview: This file displays the results of a requested export, and allows
           the user to check data.
Author(s): Jared F. McCaffree
-->

<!--#include virtual="emissions/global.asi"-->

<%
   verifyUser
   Dim objRSEmissions	' emission recordset object
   Dim oConn		' connection object
   Dim sQueryString	' the query string for the grouptable query
   Dim sPlant		' the target plant ID

   sPlant = Request.Form("plant")
   sQueryString = Request.Form("query")
   connectDB oConn
   set objRSEmissions = oConn.Execute(sQueryString)

%>

<HTML>
<HEAD>
<TITLE>Search Results</TITLE>
<LINK href="../main.css" rel="stylesheet" type="text/css">
</HEAD>
<BODY>

<% 
   printHeader "Export Search Results"

%>
<B>Please enter a file name in the text box below and verify that<BR>
the data you wish to export is correct.  To export the data to the<BR>
specified file name, press the "Export" button.</B><BR>
NOTE: The process may take several minutes to complete!<BR>
<FORM method=post action="searchexport_handler.asp">
File name: <INPUT type=text name="filename" id="filename" value=".csv"><BR>
<INPUT type=hidden id="query" name="query" value="<% = sQueryString %>">
<INPUT type=hidden id="plant" name="plant" value="<% = sPlant %>">
<INPUT type=submit value="Export">
</FORM>
<TABLE border=1>
<%
   ' if something is returned print out a table of the results
   if not objRSEmissions.EOF then
      sTableHeader = "<TH>Plant</TH><TH>Parameter</TH><TH>Test Method</TH>" _
                       & "<TH>Emission</TH><TH>Unit Number</TH><TH>Rep Number</TH>" _
                       & "<TH>Quantity</TH><TH>Import Date</TH>"
      Response.Write(sTableHeader)
      do while not objRSEmissions.EOF
         Response.Write("<TR>")
         Response.Write("<TD align=center>" & sPlant & "</TD>")
         Response.Write("<TD align=center>" & objRSEmissions.Fields("parameter") & "</TD>")
         Response.Write("<TD align=center>" & objRSEmissions.Fields("testMethod") & "</TD>")
         Response.Write("<TD align=center>" & objRSEmissions.Fields("emission") & "</TD>")
         Response.Write("<TD align=center>" & objRSEmissions.Fields("unitNumber") & "</TD>")
         Response.Write("<TD align=center>" & objRSEmissions.Fields("repNumber") & "</TD>")
         Response.Write("<TD align=center>" & objRSEmissions.Fields("emissionValue") & "</TD>")
         Response.Write("<TD align=center>" & objRSEmissions.Fields("importDate") & "</TD>")
         Response.Write("</TR>")
         objRSEmissions.MoveNext
      loop
	  Response.Write("<A href=search.asp>Return to Basic Search Page</A>")
   ' if nothing is returned let the user know
   else
      sErrMsg = "No matches were found for your query please" _
                & " <A href=search.asp>Return </A> and try again."
      Response.Write(sErrMsg)
   end if
%>

</TABLE>

<%
   printFooter 
%>

</BODY>
</HTML>

