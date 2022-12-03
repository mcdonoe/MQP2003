<!--
advsearch.asp

Overview:  This file prints a form which enables the user to select which
           test methods to view between plants
Author(s): Ron Cormier
-->


<!--#include virtual="emissions/global.asi"-->

<SCRIPT Language="VBScript">

' cmdSelectAllPlants_onclick()
' Author(s):      Jared F. McCaffree
' Overview:       When the "Select All" button is clicked this subroutine is 
'                   called. This subroutine sets the checked value of each
'                   checkbox on the page to true.
   Sub selall_onclick

      dim oFormEntry                ' for each loop value

      for each oFormEntry in forma
         if oFormEntry.type = "checkbox" then
            oFormEntry.checked = true
         end if
      next
   
   End Sub


' cmdSearchSubmit_onClick()
' Author(s):      Ron Cormier
' Overview:       This function performs form validation on the search 
'                   parameters displayed on the advanced search page.
'                   Makes sure at least one box is checked

   Sub cmdAdvancedSearch_onClick

      Dim bOneChecked
      Dim bErrTrapped
      Dim sErrMsg

      bOneChecked = false
      bErrTrapped = false
      sErrMsg = ""

      for each elem in forma
         if elem.type = "checkbox" then
            if elem.checked then
               bOneChecked = true
            end if
         end if
      next

      if not bOneChecked then
         sErrMsg = "Please select at least one test method"
         bErrTrapped = true
      end if

      if bErrTrapped = true then
         msgBox sErrMsg
      else
         forma.submit
      end if
   End Sub         

</SCRIPT>

<%
   verifyuser

   Dim oConn                    ' database connection object
   Dim objRS                    ' recordset object
   Dim sQueryString             ' string for query
   Dim sAccPlants               ' string of plants may have been selected
   Dim aAccPlants               ' array of plants may have been selected
   Dim sSelectedPlants          ' plants that were selected
   Dim sSelectedPlantID         ' comma seperated string of plant id's
   Dim aSelectedPlantID         ' array of plant id's
   Dim sTestMethods             ' comma seperated string of test methods
   Dim aTestMethods             ' array of test methods
   Dim sAllMethods              ' string of all possible test methods
   Dim aAllMethods              ' array of all possible test methods
   Dim sCommonMethods           ' string of common methods
   Dim aCommonMethods           ' array of common methods
   Dim iCnt
   Dim sPossiblePlants          ' string of plants w/ the same test method
   Dim aPossiblePlants          ' array of plants w/ the same test method
   Dim bMatchFound, bMethodFound


   ' get posted data
   sAccPlants = Request.Form("sAccPlants")
   aAccPlants = Split(sAccPlants, ", ")
   for each sPlant in aAccPlants
      if (Request.Form(sPlant) <> "") then
         sSelectedPlants = sSelectedPlants & sPlant & ","
      end if
   next
   sSelectedPlants = Left(sSelectedPlants, len(sSelectedPlants) -1)
   aSelectedPlants = Split(sSelectedPlants, ",")
 
'   Response.Write("Possible Plants: " & sAccPlants & "<BR>")
'   Response.Write("Actual Plants: " & sSelectedPlants & "<BR>")

   ' generate query string to get plant id's
   sQueryString = "SELECT ID FROM plant WHERE"
   for each sPlant in aSelectedPlants
      sQueryString = sQueryString & " plantName='" & sPlant & "' OR"
   next
   sQueryString = Left(sQueryString, Len(sQueryString)-3)


   connectDB oConn     'connect to database

   ' get string of plant id's
   sSelectedPlantID = ""
   set objRs = oConn.Execute(sQueryString)
   do while not objRS.EOF
      sSelectedPlantID = sSelectedPlantID & objRS.Fields("ID") & ","
      objRS.Movenext
   loop
   sSelectedPlantID = Left(sSelectedPlantID, Len(sSelectedPlantID)-1)
   aSelectedPlantId = Split(sSelectedPlantID, ",")



   ' generate query string to get all possible test methods
   sAllMethods = ""
   iCnt = 0
   sQueryString = "SELECT DISTINCT testMethod FROM emission"
   set objRS = oConn.Execute(sQueryString)
   do while not objRS.EOF
      'Response.Write(objRS.Fields("testMethod") & "<BR>")
      sAllMethods = sAllMethods & objRS.Fields("testMethod") & ","
      iCnt = iCnt + 1
      objRS.Movenext
   loop
   if (iCnt > 0) then
      sAllMethods = Left(sAllMethods, Len(sAllMethods)-1)      
   end if
   aAllMethods = Split(sAllMethods, ",")

   'get all test methods and then get the plant ids associated w/ each one
   '   make string of plantID's that share that method
   '   

   sCommonMethods = ""
   for each sMethod in aAllMethods
      'for each test method
      bMethodFound = true
      sPossiblePlants = ""
      sQueryString = "SELECT DISTINCT plantID FROM emission WHERE testMethod='"
      sQueryString = sQueryString & sMethod & "'"
      set objRS = oConn.Execute(sQueryString)

      ' generate a comma-seperated list of plant id's that use that method
      do while not objRS.EOF
         sPossiblePlants = sPossiblePlants & objRS.Fields("plantID") & ","
         objRS.Movenext
      loop
      if (Len(sPossiblePlants > 0)) then
         sPossiblePlants = Left(sPossiblePlants, Len(sPossiblePlants)-1)
      end if
      aPossiblePlants = Split(sPossiblePlants, ",")

      ' check to see if any of those that use that method are the same as
      ' as the ones selected on the prior page
      for each sPlant in aSelectedPlantID
         bMatchFound = false
         for each sPlant2 in aPossiblePlants
            if (sPlant = sPlant2) then
               bMatchFound = true
            end if
         next
         if (bMatchFound = false) then
            bMethodFound = false
         end if
      next

      if (bMethodFound = true) then
         sCommonMethods = sCommonMethods & sMethod & ","
      end if
   next

   disconnectDB oConn
   if Len(sCommonMethods) > 0 then
      sCommonMethods = Left(sCommonMethods, Len(sCommonMethods)-1)
   end if
   'Response.Write(sCommonMethods & "<BR>")
   aCommonMethods = Split(sCommonMethods, ",")
%>


<HTML>
<HEAD>
<TITLE>Advanced Search for Emissions - Covanta Emissions Database</TITLE>
<LINK href="../main.css" rel="stylesheet" type="text/css">
</head>

<% printHeader "Advanced Search - Select Test Methods - Step 2" %>

<BODY>
<FORM action="advsh_handler.asp" method=post id=forma name=forma>
<B>
Selected Plants: <%=sSelectedPlants%>
<BR>
Below is a list of test methods common to the plants you selected. 
Please select the emission test methods you wish to view data for.</B>
<BR>

<TABLE cellSpacing=1 cellPadding=1 border=0>
   <TR>
      <TD>
<%
   iCnt = 0
   for each sMethod in aCommonMethods
      Response.Write("<INPUT type=checkbox name='" & sMethod _
                      & "' value='" & sMethod & "'>" & sMethod _
                      & "<BR>" & vbcrlf)
      iCnt = iCnt + 1
   next
   if (iCnt = 0) then
      Response.Write("<H3>No test methods are held in commmon between" _
                     & " the selected plants.<BR>Please return to the " _
                     & "previous page and choose different plants.</H3>")
   else
%>
      <INPUT type=button value='Select All' name=selall id=selall>
      <INPUT type=hidden name='sTestMethods' value='<%=sTestMethods%>'>
      <BR><BR>
      <INPUT type=button name="cmdAdvancedSearch" id="cmdAdvancedSearch" 
       value="Submit">
      <INPUT type=hidden name=sSelPlantID value='<%=sSelectedPlantID%>'>
      <INPUT type=hidden name=sSelPlant value='<%=sSelectedPlants%>'>
      <INPUT type=hidden name=sComMeths value='<%=sCommonMethods%>'>
      <INPUT TYPE=reset>
<%
   end if
%>
      </TD>
   </TR>
</TABLE>
<BR>
</FORM>

</center>

<BR><A HREF = advsearch.asp>Back to Plant Selection</A>

<% printFooter %>


</BODY>
</HTML>