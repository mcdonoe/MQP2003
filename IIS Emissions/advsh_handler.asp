<!--
advsh_handler.asp

Overview: This file collects the information on the plants and test 
           methods selected then generates a checkbox list of emissions 
           that are common to the plants and test methods that the user
           can search for
Author(s): Ronald Cormier
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

      for each oFormEntry in form1
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
      Dim bUnitNumber
      Dim bRep1
      Dim bRep2

      bOneChecked = false
      bErrTrapped = false

      ' make sure unit number is blank or an int
      if form1.unitnumber.value <> "" then
         bUnitNumber = IsNumeric(form1.unitNumber.value)
         if bUnitNumber = false then
            sErrMsg = "Please enter an interger in the Unit Number field"
            bErrTrapped = true
         end if
      end if

      ' make sure first rep value is blank or an int
      if form1.ShowRepsFrom.value <> "" then
         bRep1 = IsNumeric(form1.ShowRepsFrom.value)
         if bRep1 = false then 
            sErrMsg = "Plase enter an integer in the first Show Repetitions Field"
            bErrTrapped = true
         end if
      end if

      ' make sure second rep value is blank or an int
      if form1.ShowRepsTo.value <> "" then
         bRep2 = IsNumeric(form1.ShowRepsTo.value)
         if bRep2 = false then
            sErrMsg = "Plase enter an integer in the second Show Repetitions Field"
            bErrTrapped = true
         end if
      end if

      ' make sure at least on emission is checked
      for each elem in form1
         if elem.type = "checkbox" then
            if elem.checked then
               bOneChecked = true
            end if
         end if
      next

      if not bOneChecked then
         sErrMsg = "Please select at least one emisson"
         bErrTrapped = true
      end if

      if bErrTrapped = true then
         msgBox sErrMsg
      else
         form1.submit
      end if
   End Sub         

</SCRIPT>



<%
   Dim objRS          ' database recordset connection object
   Dim oConn          ' database connection object
   Dim sQueryString   ' string to query dbase with
   Dim sSelPlant      ' comma seperated string of selected plants
   Dim sSelPlantID    ' comma seperated string of id's of selected plants
   Dim aSelPlantID    ' array of id's of selected plants
   Dim sComMeths      ' comma seperated string of common methods between
                      '   the selected plants
   Dim aComMeths      ' array of common methods between selected plants
   Dim sSelMeths      ' comma seperated string of methods seleced from
                      '   previous page
   Dim aSelMeths      ' array of methods selected from previous page
                      ' all selected plants have these methods in common
   Dim sEmissions     ' comma seperated string of emission common
                      '   to the plants and test methods already chosen
   Dim aEmissions     ' array of emissions common to plants and test 
                      '   methods already chosen

   ' get posted data
   sSelPlant   = Request.Form("sSelPlant")
   sSelPlantID = Request.form("sSelPlantID")
   sComMeths   = Request.Form("sComMeths")
   aSelPlantID = Split(sSelPlantID, ",")
   aComMeths = Split(sComMeths, ",")
   sSelMeths = ""
   for each comMeth in aComMeths
      if (Request.Form(comMeth) <> "") then
         sSelMeths = sSelMeths & comMeth & ","
      end if
   next  
   sSelMeths = Left(sSelMeths, Len(sSelMeths)-1)
   aSelMeths = Split(sSelMeths, ",")

   ' build query string
   sQueryString = "SELECT DISTINCT emission FROM emission WHERE"
   for each sPlantID in aSelPlantID
      sQueryString = sQueryString & " (plantID=" & sPlantID & " AND ("
      for each sMeth in aSelMeths
         sQueryString = sQueryString & "testMethod='" & sMeth & "' OR "
      next
      sQueryString = Left(sQueryString, Len(sQueryString)-4)
      sQueryString = sQueryString & ")) OR"
   next
   sQueryString = Left(sQueryString, Len(sQueryString)-3)
   sQueryString = sQueryString & " ORDER BY emission"

   connectDB oConn
   set objRS = oConn.Execute(sQueryString)

   ' get a list of possible emissions to search for 
   sEmissions = ""
   do while not objRS.EOF
      sEmissions = sEmissions & objRS.Fields("emission") & ","
      objRS.Movenext
   loop

   disconnectDB oConn

   sEmissions = Left(sEmissions, Len(sEmissions)-1)
   aEmissions = Split(sEmissions, ",")
%>


<HTML>
<HEAD>
<TITLE>Advanced Search for Emissions - Covanta Emissions Database</TITLE>
<LINK href="../main.css" rel="stylesheet" type="text/css">
</head>

<BODY>

<% printHeader "Advanced Search  -  Select Emissions - Step 3" %>


<FORM action="advres_handler.asp" method=post id=form1 name=form1>

<B>
Selected Plants: <%=sSelPlant%><BR>
Selected Methods: <%=sSelMeths%><BR>
Below is a list of emissions common to the plants and test methods 
you selected.  Please select the emissions you wish to view data for.</B>
<BR>
<TABLE>
<TR>
   <TD>Which emissions?</TD>
   <TD>
<%
   for each sEm in aEmissions
      Response.Write("<INPUT type=checkbox name='" & sEm & "' value='" _
                      & sEm & "'>" & sEm & "<BR>")
   next
%>
      <INPUT type=button value='Select All' name=selall id=selall>
   </TD>
</TR>
<TR>
   <TD>Unit Number:</TD>
   <TD><INPUT type=text name="unitNumber" size=5 id="unitNumber">*</TD>
</TR>
<TR>
   <TD>Show Repetitions:</TD>
   <TD>
      <INPUT type="text" name="showRepsFrom" id="showRepsFrom" size="5">* 
      to <INPUT type="text" name="showRepsTo" id="showRepsTo" size="5">
      <BR>
      *(Leave blank to select all.)
   </TD>
</TR>
</TABLE>
<BR><BR>
<INPUT type=button name="cmdAdvancedSearch" id="cmdAdvancedSearch" 
value="Submit">
<INPUT TYPE=reset>
<INPUT type=hidden name=sEmissions value='<%=sEmissions%>'>
<INPUT type=hidden name=sSelPlantID value='<%=sSelPlantID%>'>
<INPUT type=hidden name=sSelMeths value='<%=sSelMeths%>'>
<INPUT type=hidden name=sSelPlant value='<%=sSelPlant%>'>
</FORM>
<BR><A HREF = advsearch.asp>Back to Plant Selection</A>

<% printFooter %>

</BODY>
</HTML>