<!--
advres_handler.asp

Overview: This file collects infomation on which emission the user would
           like to view data for based on what plants and test methods
           were selected
Author(s): Ronald Cormier
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
   Dim objRS          ' database recordset connection object
   Dim oConn          ' database connection object
   Dim sQueryString   ' string to query dbase with
   Dim sSelPlant      ' comma seperated string of selected plants
   Dim sSelPlantID    ' comma seperated string of id's of selected plants
   Dim aSelPlantID    ' array of id's of selected plants
   Dim sSelMeths      ' comma seperated string of methods seleced from
                      '   previous page
   Dim aSelMeths      ' array of methods selected from previous page
                      ' all selected plants have these methods in common
   Dim sEmissions     ' comma seperated string of emission common
                      '   to the plants and test methods already chosen
   Dim aEmissions     ' array of emissions common to plants and test 
                      '   methods already chosen
   Dim sSelEm         ' comma seperated string of selected emissions
                      '   from the previous page
   Dim aSelEm         ' array of selected emissions from prev page
   Dim iRecords       ' int number of records returned
   Dim sUnitNumber    ' unit number being searched for
   Dim sShowRepsFrom  ' repetitions being searched from
   Dim sShowRepsTo    ' repetitions being searched to

   ' get posted data
   sSelPlant     = Request.Form("sSelPlant")
   sSelMeths     = Request.Form("sSelMeths")
   sUnitNumber   = Request.Form("unitNumber")
   sShowRepsFrom = Request.Form("showRepsFrom")
   sShowRepsTo   = Request.Form("showRepsTo")
   sSelPlantID   = Request.form("sSelPlantID")
   sEmissions    = Request.Form("sEmissions")
   aSelPlantID   = Split(sSelPlantID, ",")
   aSelMeths     = Split(sSelMeths, ",")
   aEmissions    = Split(sEmissions, ",")
   sSelEm = ""
   for each sEm in aEmissions
      if (Request.Form(sEm) <> "") then
         sSelEm = sSelEm & sEm & ","
      end if
   next
   sSelEm = Left(sSelEm, Len(sSelEm)-1)
   aSelEm = Split(sSelEm, ",")

   ' build query string
   sQueryString = "SELECT plant.plantName AS name,parameter,testMethod,"
   sQueryString = sQueryString & "emission,unitNumber,repNumber,"
   sQueryString = sQueryString & "emissionValue,importDate"
   sQueryString = sQueryString & " FROM emission INNER JOIN plant ON "
   sQueryString = sQueryString & "emission.plantID=plant.ID WHERE ("
   for each sPlantID in aSelPlantID
      sQueryString = sQueryString & " (plantID=" & sPlantID & " AND ("
      for each sMeth in aSelMeths
         sQueryString = sQueryString & "testMethod='" & sMeth & "' OR "
      next
      sQueryString = Left(sQueryString, Len(sQueryString)-4)
      sQueryString = sQueryString & ") AND ("
      for each sEm in aSelEm
         sQueryString = sQueryString & "emission='" & sEm & "' OR "
      next
      sQueryString = Left(sQueryString, Len(sQueryString)-4)
      sQueryString = sQueryString & ")) OR"
   next

   sQueryString = Left(sQueryString, Len(sQueryString)-3)
   sQueryString = sQueryString & ")"

   sQueryString = sQueryString & " AND "
   if sUnitNumber <> "" then
      sQueryString = sQueryString & "(unitNumber=" & sUnitNumber & ") AND "
   end if

   if sShowRepsFrom <> "" and sShowRepsTo <> "" then
      sQueryString = sQueryString & "(repNumber >= " & sShowRepsFrom _
                     & " AND repNumber <= " & sShowRepsTo & ") AND "
   end if
   sQueryString = Left(sQueryString, Len(sQueryString)-5)
   sQueryString = sQueryString & " ORDER BY plantID,emission, parameter,"
   sQueryString = sQueryString & "testMethod, unitNumber, repNumber"

   connectDB oConn

   'Response.Write(sQueryString)
   set objRS = oConn.Execute(sQueryString)
%>


<HTML>
<HEAD>
<TITLE>Advanced Search Results</TITLE>
<LINK href="../main.css" rel="stylesheet" type="text/css">
<BODY>

<% printHeader "Advanced Search - Results" %>
<FORM method=post action="search_export.asp">
<INPUT type=hidden id="query" name="query" value="<% = sQueryString %>">
<INPUT type=hidden id="plant" name="plant" value="<% = sSelPlant %>">
<INPUT type=submit value="Export Search Results">
</FORM>
<B>
Selected Plants: <%=sSelPlant%><BR>
Selected Methods: <%=sSelMeths%><BR>
Selected Emissions: <%=sSelEm%><BR>
Selected Unit Number: <%=sUnitNumber%><BR>
Selected Repetitions: <%=sShowRepsFrom%> to <%=sShowRepsTo%><BR>
</B>
<TABLE border=1 id="table1" name="table1">
<%
   ' if something is returned print out a table of the results
   iRecords = 0
   if not objRS.EOF then
      sTableHeader = "<TH><A HREF = '#' onclick ='return sortTable(0)'>Plant</A></TH><TH>" _
                       & "<A HREF = '#' onclick = 'return sortTable(1)'>Paramter</A></TH><TH>" _
		       & "<A HREF = '#' onclick = 'return sortTable(2)'>Test Method</A></TH><TH>" _
                       & "<A HREF = '#' onclick ='return sortTable(3)'>" _
                       & "Emission</A></TH><TH><A HREF ='#'" _
                       & "onclick='return sortTable(4)'>Unit Number</A></TH>" _
                       & "<TH><A HREF ='javascript'" _
                       & "onclick='sortTable(5); return false;'>Rep Number</A> </TH>" _
                       & "<TH><A HREF ='javascript'" _
                       & "onclick='sortTable(6); return false;'>Quantity</TH><TH>" _
                       & "<A HREF ='#'" _
                       & "onclick='return sortTable(7)'>Import Date</A></TH>"
      Response.Write(sTableHeader)
      do while not objRS.EOF
         Response.Write("<TR>")
         Response.Write("<TD align=center>" & objRS.Fields("name") & "</TD>")
         Response.Write("<TD align=center>" & objRS.Fields("parameter") & "</TD>")
         Response.Write("<TD align=center>" & objRS.Fields("testMethod") & "</TD>")
         Response.Write("<TD align=center>" & objRS.Fields("emission") & "</TD>")
         Response.Write("<TD align=center>" & objRS.Fields("unitNumber") & "</TD>")
         Response.Write("<TD align=center>" & objRS.Fields("repNumber") & "</TD>")
         Response.Write("<TD align=center>" & objRS.Fields("emissionValue") & "</TD>")
         Response.Write("<TD align=center>" & objRS.Fields("importDate") & "</TD>")
         Response.Write("</TR>" & vbcrlf)
         iRecords = iRecords + 1
         objRS.MoveNext
      loop
   ' if nothing is returned let the user know
   else
      sErrMsg = "No matches were found for your query please" _
                & " <A href=search.asp>Return </A> and try again."
      Response.Write(sErrMsg)
   end if

   disconnectDB oConn
%>
</TABLE>
<%=iRecords%> record(s)<BR>
<A href=advsearch.asp>Return to Advanced Search Page</A>
<FORM method=post action="search_export.asp">
<INPUT type=hidden id="query" name="query" value="<% = sQueryString %>">
<INPUT type=hidden id="plant" name="plant" value="<% = sSelPlant %>">
<INPUT type=submit value="Export Search Results">
</FORM>
<% printFooter %>
</BODY>
</HTML></HTML>