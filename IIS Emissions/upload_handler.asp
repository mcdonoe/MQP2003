<!--
upload_handler.asp

Overview:  This file handles the upload of the file as specified in the 
             importdata.asp file.
NOTE:      You must have VBScript 5.0 installed on your web server in order for
             the file uploading to work, specifically upload.asp.
Author(s): Jared F. McCaffree
-->

<!--#include virtual="emissions/global.asi"-->
<!--#include file="upload.asp" -->


<HTML>
<HEAD>
<%
   verifyUser
%>

<TITLE>Import Data Step 2: Import Uploaded File</TITLE>
<LINK href="../main.css" rel="stylesheet" type="text/css">

<BODY>

   <% printHeader "Import Data Step 2: Import Uploaded File" %>

<B>The file has been uploaded.  Below is the file in raw comma <BR>
delimited format.  Please make sure this output matches the file <BR>
that is to be imported.<BR>
<BR>
To finish the upload process press the Import Data button below<BR>
This will add the data to the main database table.<BR>
</B>

<%

   Dim Uploader       ' file uploader object
   Dim File           ' uploaded file object
   Dim sFileName       ' filename of the output file
   Dim objFSO          ' file system object
   Dim objTextFile     ' output file object
   Dim sInString       ' input buffer for display output file to screen
   Dim sPathToUpload   ' path to output file without filename
   Dim sPath           ' absolute path to output file

   ' Create the file uploader object
   Set Uploader = New FileUploader

   ' This starts the upload process
   Uploader.Upload()

   ' Save the file to disk
   if Uploader.Files.Count = 0 then
      Response.Write "File(s) not uploaded <br>"
      Response.Write "File count: " & Uploader.Files.Count & "<br>"
   else
      For Each File In Uploader.Files.Items
         File.SaveToDisk sPathToUpload

         Set objFSO = CreateObject("Scripting.FileSystemObject")
         sFileName = File.FileName
      Next
      sPath = sPathToUpload & sFileName

      ' debugging
      'Response.Write("sPath: " & sPath & "<br>")
      'Response.Write("sFileName: " & sFileName & "<br>")
      'Response.Write("file count: " & Uploader.Files.Count & "<br>")
   end if

%>

<FORM action="importdata_handler.asp" method=POST>
<input type=hidden name="fileName" value="<% =sFileName %>">
<input type=hidden name="plant" value="<% =Uploader.Form("plants") %>">

<input type=submit value="Import Data">
</FORM>

<%
   ' Print out the data written to the file
   Response.Write("<B>Import to plant: " & Uploader.Form("plants") & "</B><BR>")
   Response.Write("<B>Data to be Imported: </B>")
   Response.Write("<pre>")
   Set objTextFile = objFSO.OpenTextFile(sPath)
   Do While Not objTextFile.AtEndOfStream
      sInString = objTextFile.Readline
      Response.Write(sInString & vbcrlf)
   Loop
   Response.Write("</pre>")
%>
<BR>

   <% printFooter %>




</BODY>
</HTML>