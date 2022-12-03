<!--
importdata.asp

Overview:  HTML form for upload
Author(s): Ronald Cormier (rcormier@wpi.edu)
           Jared McCaffree (jared@wpi.edu)
-->

<!--#include virtual="emissions/global.asi"-->

<%
   verifyUser   	'verify user is logged in

    Dim sPlantString	'make the HTML code for the plant radio boxes
    sPlantString = printPlants("-r", "")
%>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>Import Data Step 1: Upload Data</TITLE>
</HEAD>
<LINK href="../main.css" rel="stylesheet" type="text/css">
<BODY>
<H3>Import Data Step 1: Upload Data</H3><HR>
<B>The first step of the data import process is uploading the data to the server.<BR>
The data must be in comma delimited format; for example: newdata.csv<BR>
To save a file in Microsoft Excel in comma deliminated format select the File menu, <BR>
followed by Save As.  Type in the filename, and select "CSV (Comma Delimited) (*csv)"<BR>
from the Save as type: list below the file name field.
</B><BR>
<a href="./">Back to Menu</a>
<BR>
<BR>
<FORM METHOD="POST" ENCTYPE="multipart/form-data" ACTION="upload_handler.asp">

	<TABLE BORDER=0>
	<TR><TD><B>Select data to upload:</B><BR></TD>
            <TD><INPUT TYPE=FILE name="File1"></TD></TR>
        <TR><TD><B>Select plant:</B></TD>
        <TD><% =sPlantString %></TD></TR>
        <TR><TD><INPUT type = submit value="Continue to Step 2"></TD></TR>
	</TABLE>
</FORM>
</BODY>
</HTML>
