<!--
search_handler.asp

Overview: This file queries the database based on the information passed
          from the search.asp file.  It then sorts and displays the data
          in a table.
Author(s): Jared F. McCaffree
           Ron Cormier
-->

<!--#include virtual="emissions/global.asi"-->


<SCRIPT Language = "JavaScript">
currentCol = 0

   function sortTable(sortBy){
    
      var tableSource = document.all('table1');

      var numRows = tableSource.rows.length - 1;
      var numCols = tableSource.rows(0).cells.length;
      var tableArray = new Array(numRows)
      for (i=0; i<numRows; i++) {
          tableArray[i] = new Array(numCols)
          
          for (j=0; j<numCols; j++)
              tableArray[i][j] = tableSource.rows(i+1).cells(j).innerText
      }

      //alert(tableArray[0][0]);
      //alert("table has rows: " +numRows + " and cols: " + numCols);

      currentCol = sortBy
      if (sortBy == 7)
         tableArray.sort(CompareDate);
      else if (sortBy == 0 || sortBy == 1 || sortBy == 2 || sortBy == 3){
	tableArray.sort(CompareAlpha);
        //alert(sortBy);
      }
      else
      tableArray.sort(CompareInteger);

      for (i=0; i<numRows; i++) {
          for (j=0; j<numCols; j++) {
               tableSource.rows(i+1).cells(j).innerText = tableArray[i][j];
          }
      }

   }

function CompareAlpha(a, b) {
	if (a[currentCol] < b[currentCol]) { return -1; }
	if (a[currentCol] > b[currentCol]) { return 1; }
	return 0;
}
   function CompareInteger(a,b) {
    numA = a[currentCol]
    //alert(numA);
    numB = b[currentCol]

  
    return numA - numB;
     
   }        

function CompareDate(a, b) {
	// this one works with date formats conforming to Javascript specifications, e.g. m/d/yyyy
	datA = new Date(a[currentCol]);
	datB = new Date(b[currentCol]);
	if (datA < datB) { return -1; }
	else {
		if (datA > datB) { return 1; }
		else { return 0; }
	}
}

</SCRIPT>

<% 
   verifyUser

   Dim sQueryString	' the query string for the grouptable query
   Dim oConn		' connection object
   Dim objRS		' grouptable recordset object
   Dim objRSEmissions	' emission recordset object
   Dim sPlant          	' plant being searched for
   Dim iPlantID         ' the ID of the plant being searched for
   Dim sUnitNumber      ' unit number being searched for
   Dim sTestMethod      ' test method being searched for
   Dim sShowRepsFrom    ' repetitions being searched from
   Dim sShowRepsTo      ' repetitions being searched to
   Dim sEmissions       ' string of emissions that may have been searched
   Dim aEmissions       ' array of emissions that may have been searched
   Dim sActualEm        ' string of emissions that were actually searched for
   Dim aActualEm        ' array of emissions that were actually searched for
   Dim iRecords         ' integer number of records returned
   Dim sTestMethods     ' string of test methods that may have been searched
   Dim aTestMethods     ' array of test methods that may have been searched
   Dim sActualMeth      ' string of test methods actually searched for
   Dim aActualMeth      ' array of test methods actually searched for


   ' Record the form data
   sPlant        = Request.Form("plant")
   sUnitNumber   = Request.Form("unitNumber")
   sTestMethod   = Request.Form("testMethod")
   sShowRepsFrom = Request.Form("showRepsFrom")
   sShowRepsTo   = Request.Form("showRepsTo")
   sEmissions    = Request.Form("sEmissions")
   sTestMethods  = Request.Form("sTestMethod")

   ' generate a comma sepereated string of all emissions searched for
   sActualEm = ""
   aEmissions = Split(sEmissions, ",")
   for each emission in aEmissions
      if (Request.Form(emission) <> "") then
         sActualEm = sActualEm & Request.Form(emission) & ","
      end if
   next

   sActualEm = left(sActualEm, len(sActualEm)-1)


   ' generate a comma sepereated string of all methods searched for
   sActualMeth = ""
   aTestMethods = Split(sTestMethods, ",")
   for each method in aTestMethods
      if (Request.Form(method) <> "") then
         sActualMeth = sActualMeth & Request.Form(method) & ","
      end if
   next

   sActualMeth = left(sActualMeth, len(sActualMeth)-1)



   ' set the query string to select the plant's ID and connect to the database
   sQueryString = "SELECT DISTINCT * FROM plant WHERE plantName='" & sPlant & "'"

   connectDB oConn
   ' retrieve the ID of the plant
   set objRS = oConn.Execute(sQueryString)
   iPlantID = objRS.Fields("ID")   

   ' compile searching SQL query string
   sQueryString = "SELECT * FROM emission WHERE "
   if sPlant <> "" then
      sQueryString = sQueryString & "plantID='" & iPlantID & "'"
   end if

   ' include the test methods that were actually seached for in sql query
   aActualMeth = Split(sActualMeth, ",")
   sQueryString = sQueryString & " AND ("
   for each method in aActualMeth
      sQueryString = sQueryString & " testMethod='" & method & "' OR"
   next
   sQueryString = left(sQueryString, len(sQueryString)-2)
   sQueryString = sQueryString & ")"


   ' include the emission that were actually seached for in sql query
   aActualEm = Split(sActualEm, ",")
   sQueryString = sQueryString & " AND ("
   for each emiss in aActualEm
      sQueryString = sQueryString & " emission='" & emiss & "' OR"
   next
   sQueryString = left(sQueryString, len(sQueryString)-2)
   sQueryString = sQueryString & ")"


   if sUnitNumber <> "" then
      sQueryString = sQueryString & " AND unitNumber=" & sUnitNumber
   end if

   if sTestMethod <> "" then
      sQueryString = sQueryString & " AND testMethod='" & sTestMethod & "'"
   end if

   if sShowRepsFrom <> "" and sShowRepsTo <> "" then
      sQueryString = sQueryString & " AND repNumber >= " & sShowRepsFrom _
                     & " AND repNumber <= " & sShowRepsTo
   end if

   
   sQueryString = sQueryString & " ORDER BY emission,parameter,testMethod,unitNumber,repNumber"
   ' get the resulting recordset
   set objRSEmissions = oConn.Execute(sQueryString)

%>


<HTML>
<HEAD>
<TITLE>Search Results</TITLE>
<LINK href="../main.css" rel="stylesheet" type="text/css">
<BODY>

<% printHeader "Search Results" %>

<FORM method=post action="search_export.asp">
<INPUT type=hidden id="query" name="query" value="<% = sQueryString %>">
<INPUT type=hidden id="plant" name="plant" value="<% = sPlant %>">
<INPUT type=submit value="Export Search Results">
</FORM>
<B>
Selected Plants: <%=sPlant%><BR>
Selected Methods: <%=sActualMeth%><BR>
Selected Emissions: <%=sActualEm%><BR>
Selected Unit Number: <%=sUnitNumber%><BR>
Selected Repetitions: <%=sShowRepsFrom%> to <%=sShowRepsTo%><BR>
</B>
<TABLE border=1 id="table1" name="table1">
<%
   ' if something is returned print out a table of the results
   iRecords = 0
   if not objRSEmissions.EOF then
      sTableHeader = "<TH><A HREF = '#' onclick ='return sortTable(0)'>Plant</A></TH>" _
                   & "<TH><A HREF = '#' onclick ='return sortTable(1)'>Paramter</A></TH>" _
                   & "<TH><A HREF = '#' onclick ='return sortTable(2)'>Test Method</A></TH>" _
                   & "<TH><A HREF = '#' onclick ='return sortTable(3)'>Emission</A></TH>" _
                   & "<TH><A HREF = '#' onclick ='return sortTable(4)'>Unit Number</A></TH>" _
                   & "<TH><A HREF = '#' onclick ='return sortTable(5)'>Rep Number</A></TH>" _
                   & "<TH><A HREF = '#' onclick ='return sortTable(6)'>Quantity</A></TH>" _
                   & "<TH><A HREF = '#' onclick ='return sortTable(7)'>Import Date</A></TH>"
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
         Response.Write("</TR>" & vbcrlf)
         iRecords = iRecords + 1
         objRSEmissions.MoveNext
      loop

   ' if nothing is returned let the user know
   else
      sErrMsg = "No matches were found for your query please" _
                & " <A href=search.asp>Return </A> and try again."
      Response.Write(sErrMsg)
   end if
%>
</TABLE>
<%=iRecords%> record(s)
<BR>
<A href=search.asp>Return to Basic Search Page</A>
<FORM method=post action="search_export.asp">
<INPUT type=hidden id="query" name="query" value="<% = sQueryString %>">
<INPUT type=hidden id="plant" name="plant" value="<% = sPlant %>">
<INPUT type=submit value="Export Search Results">
</FORM>
<%  printFooter %>
</BODY>
</HTML>