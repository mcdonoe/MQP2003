<!--
search.asp

Overview: This file prints a form which enables the user to search through the
          database of emission data.
Author(s): Jared F. McCaffree
           Ron Cormier
-->

<!--#include virtual="emissions/global.asi"-->

<SCRIPT Language = "VBSCRIPT">
   Sub cmdBasicSearch_onClick
      Dim boolUnitNumber
      Dim boolTestMethod
      Dim boolRep1
      Dim boolRep2
      Dim errMsg
      Dim errTrapped      
      Dim oneEmChecked
      Dim oneMethodChecked

      oneEmChecked = false
      oneMethodChecked = false
      errTrapped = false

      ' make sure unit number is blank or an int
      if form1.unitnumber.value <> "" then
         boolUnitNumber = IsNumeric(form1.unitNumber.value)
         if boolUnitNumber = false then
            errMsg = "Please enter an interger in the Unit Number field"
            errTrapped = true
         end if
      end if

      ' make sure first rep value is blank or an int
      if form1.ShowRepsFrom.value <> "" then
         boolRep1 = IsNumeric(form1.ShowRepsFrom.value)
         if boolRep1 = false then 
            errMsg = "Plase enter an integer in the first Show Repetitions Field"
            errTrapped = true
         end if
      end if

      ' make sure second rep value is blank or an int
      if form1.ShowRepsTo.value <> "" then
         boolRep2 = IsNumeric(form1.ShowRepsTo.value)
         if boolRep2 = false then
            errMsg = "Plase enter an integer in the second Show Repetitions Field"
            errTrapped = true
         end if
      end if

      ' make sure one emission is checked
      for each formEntry in form1
         if formEntry.type = "checkbox" then
            if formEntry.Checked then
               if formEntry.id = "emission" then
                  oneEmChecked=true
               end if
            end if
         end if
      next
      if not oneEmChecked then
         errMsg = "Please select at least one emission."
         errTrapped = true
      end if

      for each formEntry in form1
         if formEntry.type = "checkbox" then
            if formEntry.Checked then
               if formEntry.id = "method" then
                  oneMethodChecked=true
               end if
            end if
         end if
      next
      if not oneMethodChecked then
         errMsg = "Please select at least one test method."
         errTrapped = true
      end if

      if errTrapped = true then
         msgBox errMsg
      else 
         Form1.submit
      end if

	
   End Sub

' cmdSelectAllEms_onclick()
' Author(s):      Ron Cormier
' Overview:       When the "Select All" button is clicked this subroutine is called.
'                   This subroutine sets the checked value of each checkbox on the page
'                   with an id of 'emission' to true.

   Sub cmdSelectAllEms_onclick

      dim oFormEntry                ' for each loop value

      for each oFormEntry in form1
         if oFormEntry.type = "checkbox" then
            if oFormEntry.id = "emission" then
               oFormEntry.checked = true
            end if
         end if
      next
   
   End Sub

' cmdSelectAllMethods_onclick()
' Author(s):      Ron Cormier
' Overview:       When the "Select All" button is clicked this subroutine is called.
'                   This subroutine sets the checked value of each checkbox on the page
'                   with an id of 'method' to true.

   Sub cmdSelectAllMethods_onclick

      dim oFormEntry                ' for each loop value

      for each oFormEntry in form1
         if oFormEntry.type = "checkbox" then
            if oFormEntry.id = "method" then
               oFormEntry.checked = true
            end if
         end if
      next
   
   End Sub
</SCRIPT>

<%
   verifyUser
   
   Dim iUserID		' the user's ID (stored in session global variable)
   Dim sQueryString	' the query string for the grouptable query
   Dim oConn		' connection object
   Dim objRS		' grouptable recordset object
   Dim objRSPlants	' plant recordset object
   Dim sPlantString	' string of plants retrieved from usertable query
   Dim asPlants         ' array of plants to print out
   Dim sEmissions       ' string of emissions that were measured
   Dim aEmissions	' array of emissions that were measured
   Dim sTestMethod      ' string of test methods
   Dim aTestMethod      ' array of test methods

   iUserID = Session("UserID")		' retrieve user's info
   sQueryString = "SELECT * FROM usertable WHERE userID =" & iUserID 
   
   ' connect to the database and execute the query string
   connectDB oConn

   set objRS = oConn.Execute(sQueryString)
   sPlantString = objRS.Fields("plant")
   ' trim and split the plant string
   sPlantString = trim(sPlantString)
   asPlants = Split(sPlantString,", ")

   ' get the emissions that were measured
   sQueryString = "SELECT DISTINCT emission FROM emission ORDER BY emission"
   set objRS = oConn.Execute(sQueryString)
   sEmissions = ""
   do while not objRS.EOF
      sEmissions = sEmissions & trim(objRS.Fields("emission")) & ","
      objRS.MoveNext
   loop

   'get the test methods that were measured for drop down list
   sQueryString = "SELECT DISTINCT testMethod FROM emission ORDER BY testMethod"
   set objRS = oConn.Execute(sQueryString)
   sTestMethod = ""
   do while not objRS.EOF
      sTestMethod = sTestMethod & trim(objRS.Fields("testMethod")) & ","
      objRS.MoveNext
   loop

   disconnectDB oConn
   sEmissions = left(sEmissions, len(sEmissions) - 1)
   sTestMethod = left(sTestMethod, len(sTestMethod) - 1)
   'Response.write(sTestMethod & "<BR>")
   aEmissions = Split(sEmissions, ",")
   aTestMethod = Split(sTestMethod, ",")

%>



<HTML>
<HEAD>
<TITLE>Search Emissions</TITLE>
<LINK href="../main.css" rel="stylesheet" type="text/css">

<BODY>
<% printHeader "Basic Search" %>
<b>Please fill out the form below and press the Search button to submit 
    your query.</b><br>

<TABLE align=right> <tr><td><a href="advsearch.asp">Advanced Search<BR></TD></TR></TABLE>
<BR>

<FORM method=post name=form1 action="search_handler.asp">
<TABLE border="0">
   <TR>
     <TD>Search Plant:</TD>
     <TD><SELECT name="plant">
<%
   for each plant in asPlants
       Response.Write("<OPTION>" & plant & "</OPTION> & vbcrlf")
   next
%>
      </SELECT></TD>
   </TR>
     <TD>Test Method:</TD>
     <TD>
<%
   for each method in aTestMethod
      Response.Write("<INPUT TYPE=checkbox name='" & method _
                      & "' value='"& method &"' id=method>" & method & "<BR>" _
                      & vbcrlf)
   next
%>
     <INPUT type="button" value="Select All" id="cmdSelectAllMethods" 
          name="cmdSelectAllMethods">
     <INPUT type=hidden name=sTestMethod value='<%=sTestMethod%>'
     </TD>
   </TR>
   <TR>
     <TD>Which emissions?</TD>
     <TD>
<%
   for each emission in aEmissions
      Response.Write("<INPUT TYPE=checkbox name='" & emission _
                      & "' value='"& emission &"' id=emission>" _
                      & emission & "<BR>" & vbcrlf)
   next
%>
     <INPUT type="button" value="Select All" id="cmdSelectAllEms" 
          name="cmdSelectAllEms">
     <INPUT TYPE=hidden name=sEmissions value='<%=sEmissions%>'>
     </TD>
   </TR>
   <TR>
     <TD>Unit Number:</TD>
     <TD><INPUT type=text name="unitNumber" size=5 id="unitNumber">*
     </TD>
   </TR>
   <TR>
     <TD valign=top>Show Repetitions: </TD>
     <TD><INPUT type="text" name="showRepsFrom" id="showRepsFrom" size="5">* 
         to <INPUT type="text" name="showRepsTo" id="showRepsTo" size="5">
        <BR>
        *(Leave blank to select all.)</TD>
   </TR>
</TABLE><BR>
<INPUT type=button name="cmdBasicSearch" id="cmdBasicSearch" value="Search">
<INPUT type=reset>

</FORM>

   <% printFooter %>

</BODY>
</HTML>