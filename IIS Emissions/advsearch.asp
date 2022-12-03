<!--
advsearch.asp

Overview: This file prints a form which enables the user to search through the
          database of emission data.  It allows the user to select a multiple
          number of plants and emissions for comparison.
Author(s): Matthew M. Barrett, Eric P. McDonough
-->


<!--#include virtual="emissions/global.asi"-->

<%
   verifyUser

   Dim allPlants

   Dim oConn   ' Connection object
   Dim objRS   ' Recordset object (result of SQL SELECT statement)	  
   iUserID = Session("UserID")		' retrieve user's info
	  
   connectDB oConn

   'get plants user has access to bu querying the user table with the 
   'user' current ID	

   Set objRS = oConn.Execute("SELECT * FROM usertable WHERE userID = " &_
               iUserID)
	  
   sPlantString = objRS.Fields("plant")
   
   'trim and split the plant string
   sPlantString = trim(sPlantString)
   asPlants = Split(sPlantString,", ")         

   disconnectDB oConn
%>

<HTML>
<HEAD>


<SCRIPT Language="VBScript">

' cmdSubmitButton_onclick()
' Author(s):      Ronald Cormier
' Overview:       This subroutine is called when the Administrator presses the 
'                 submit button of the Advanced Search form.  It checks to make
'                 sure the Administrator checks at least one plant from the 
'                 Accessible Plants list.
' Preconditions:  The Administrator clicks the submit form button after filling 
'                 out the Advanced Search - Step 1 form.
' Postconditions: If at least one checkbox has been checked, the form is 
'                 submitted.  If not, an error is displayed.

   Sub cmdSubmitButton_onclick

      dim oneChecked	' Boolean for checking validity of plant selection
      dim errMsg        ' String of the error message to be displayed
      dim errTrapped    ' Boolean, set true if an error occurrs

      oneChecked=false
      errTrapped=false

      for each formEntry in plants
         if formEntry.type = "checkbox" then
            if formEntry.Checked then
               oneChecked=true
            end if
         end if
      next
      if not oneChecked then
         errMsg = "Please select at least one accessible plant."
         errTrapped = true
      end if
      if errTrapped or not oneChecked then
         MsgBox errMsg
      else
         plants.submit
      end if
   End Sub
</SCRIPT>


<TITLE>Advanced Search for Emissions - Covanta Emissions Database</TITLE>
<LINK href="../main.css" rel="stylesheet" type="text/css">

<BODY>

<% printHeader "Advanced Search - Select Plants - Step 1" %>

<B>This search method will allow you to compare emissions test data from the 
following plants.  Please select the plants you wish to search from.</B>
<BR>
<FORM action = "advsearch_handler.asp" method=POST id=plants name=plants>
<TABLE cellSpacing=1 cellPadding=1 border=0>
   <TR>
      <TD>
<%
   for each sPlant in asPlants
	Response.Write("<INPUT TYPE=checkbox name='" & sPlant & "' value=" _
                        & sPlant & ">" & sPlant & "<BR>" & vbcrlf)
   next
%>
      </TD>
   </TR>
</TABLE>
<BR>
<INPUT TYPE=hidden name=sAccPlants value='<%=sPlantString%>'>
<INPUT TYPE="button" value="Next" id="cmdSubmitButton" name="cmdSubmitButton">
</FORM>
   <% printFooter %>
</HTML>