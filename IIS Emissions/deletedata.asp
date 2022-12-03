<!--
deletedata.asp

Overview: This file displays a listing of all the plants to potentially 
          remove data from from the emissions table.
Author(s): Jared F. McCaffree
           Ronald Cormier
-->

<!--#include virtual="emissions/global.asi"-->
<HTML>
<HEAD>
<%
   verifyUser
   
   Dim sPlantString	'make the HTML code for the plant radio boxes
   sPlantString = printPlants("-r", "")
%>

<TITLE>Delete Data</TITLE>
<LINK href="../main.css" rel="stylesheet" type="text/css">

<BODY>

<H3>Delete Data Step 1: Select Plant</H3><HR>
<B>The first step of the data delete process is selecting the plant 
from which to remove the data.<BR>
</B><BR>
<a href="./">Back to Menu</a>
<BR>
<BR>
<FORM METHOD="POST" ACTION="deletedata_handler.asp">
     <TABLE BORDER=0>
        <TR><TD><B>Select plant:</B></TD>
        <TD><% =sPlantString %></TD></TR>
        <TR><TD colspan=2><INPUT type = submit value="Continue to Step 2">
        </TD></TR>
     </TABLE>
</FORM>




</BODY>
</HTML>