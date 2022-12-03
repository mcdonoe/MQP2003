<!--
search_export.asp

Overview: This file exports the results from an SQL query received by the search
           page.  It prints a comma separated database containing the information 
           returned from querying the emissions table.
Author(s): Jared F. McCaffree
-->

<!--#include virtual="emissions/global.asi"-->

<%
   verifyUser
   Dim objRSEmissions	' emission recordset object
   Dim oConn		' connection object
   Dim sQueryString	' the query string for the grouptable query
   Dim sPlant		' the target plant ID
   Dim aEmissions(20)	' array of emissions searched for
   Dim iNumEmissions	' number of emission in aEmissions
   Dim aParameters(20)	' array of parameters searched for
   Dim iNumParameters	' number of parameters in aParameters
   Dim bFoundOne	' boolean search switch
   Dim aUnits(30)	' array of unit numbers
   Dim iNumUnits	' number of units in aUnits
   Dim aReps(30)	' array of rep numbers
   Dim iNumRepts	' number of res in aReps
   Dim iClearIndex	' counter for units
   Dim iNumPlants	' number of plants in asPlant

   ' get the plants searched for and the query string and connect to the database
   sPlant = Request.Form("plant")
   asPlant = split(sPlant,",")
   sQueryString = Request.Form("query")
   connectDB oConn
   set objRSEmissions = oConn.Execute(sQueryString)

   ' put all the different Emission types in one array
   iNumEmissions = 0
   Do while not objRSEmissions.EOF
      bFoundOne = false
      for each item in aEmissions
         if objRSEmissions.Fields("emission") = item then
            bFoundOne = true
         end if
      next
      if not bFoundOne then
         aEmissions(iNumEmissions) = objRSEmissions.Fields("emission")
         iNumEmissions = iNumEmissions + 1
      end if
      objRSEmissions.MoveNext
   Loop
   objRSEmissions.MoveFirst

   ' put all the different Parameters in one array
   iNumUnits = 0
   Do while not objRSEmissions.EOF
      bFoundOne = false
      for each unit in aUnits
         if objRSEmissions.Fields("unitNumber") = unit then
            bFoundOne = true
         end if
      next
      if not bFoundOne then
         aUnits(iNumUnits) = objRSEmissions.Fields("unitNumber")
         iNumUnits = iNumUnits + 1
      end if
      objRSEmissions.MoveNext
   Loop



   ' Declare variables for the File System Object and the File to be accessed.
   Dim objFSO, objTextFile
   ' Create an instance of the the File System Object and assign it to objFSO.
   Set objFSO = CreateObject("Scripting.FileSystemObject")
   ' Open the file
   sFileName = Server.MapPath("export") & "\" & Request.Form("filename")
   Set objTextFile = objFSO.CreateTextFile(sFileName)

   ' write the first line of the file including all of the emissions
   objTextFile.Write ",,,,"
   for each item in aEmissions
      if item <> "" then
         objTextFile.Write (item & ",")
      end if
   next
   objTextFile.Write(vbcrlf)
   ' write the start of the second line
   objTextFile.Write("Test Site,Unit #,Rep #,Parameter,")

   ' get number of plants; used for distinguishing between advanced
   '   and basic searches
   iPlant=0
   for each plant in asPlant
      iPlant = iPlant+1
   next

' For each plant that is searched for
for each plant in asPlant

   ' write the plant's name (3 character code)
   objTextFile.Write(vbcrlf)
   objTextFile.Write(plant)

   ' for each unit in that plant, print out the emission values for each
   '   parameter and for each test method
   for each unit in aUnits
    if unit <> "" then
      iNumUnits = 0
      objRSEmissions.MoveFirst
      ' compile list of all parameters in that unit
      Do while not objRSEmissions.EOF
         bFoundOne = false
         if objRSEmissions.Fields("unitNumber") = unit then
            for each item in aParameters
               if objRSEmissions.Fields("parameter") = item then
                  bFoundOne = true
               end if
               ' if it's a simple search there's only one plant, so compare
               if iPlant = 1 then 
                  if plant <> sPlant then
                     bFoundOne = true
                  end if
               else   ' if it's not a simple search get the field value and compare
                  if objRSEmissions.Fields("name") <> plant then
                     bFoundOne = true
                  end if
               end if
            next
            if not bFoundOne then
               aParameters(iNumParameters) = objRSEmissions.Fields("parameter")
               iNumParameters = iNumParameters + 1
            end if
         end if
         objRSEmissions.MoveNext
      Loop
      objTextFile.Write(",,," & item & "," & vbcrlf)
      iClearIndex=0
      for each item in aParameters
         ' print out the unit
         if item <> "" then
            if iClearIndex = 0 then
               objTextFile.Write("," & unit & ",," & item & vbcrlf)
            else
               objTextFile.Write(",,," & item & vbcrlf)
            end if
         end if
         ' get all the reps for that unit
         iNumReps = 0
         objRSEmissions.MoveFirst
         Do while not objRSEmissions.EOF
            bFoundOne = false
            for each rep in aReps
               if objRSEmissions.Fields("repNumber") = rep then
                  bFoundOne = true
               end if
               if objRSEmissions.Fields("parameter") <> item then
                  bFoundOne = true
               end if
               if objRSEmissions.Fields("unitNumber") <> unit then
                  bFoundOne = true
               end if
               if iPlant = 1 then
                  if plant <> sPlant then
                     bFoundOne = true
                  end if
               else
                  if objRSEmissions.Fields("name") <> plant then
                     bFoundOne = true
                  end if
               end if
            next
            if not bFoundOne then
               aReps(iNumReps) = objRSEmissions.Fields("repNumber")
               iNumReps = iNumReps + 1
            end if
            objRSEmissions.MoveNext
         Loop
         iClear = 0
         ' get all the unique repetitions and write them to the file
         for each rep in aReps
            if rep <> "" then
               objTextFile.Write(",," & rep & ",,")
               for each emission in aEmissions
                 if emission <> "" then
                  objRSEmissions.MoveFirst
                  if iPlant = 1 then
                     Do while not objRSEmissions.EOF
                        if objRSEmissions.Fields("repNumber") = rep and _
                          objRSEmissions.Fields("parameter") = item and _
                          objRSEmissions.Fields("unitNumber") = unit and _
                          sPlant = plant and _                   
                          objRSEmissions.Fields("emission") = emission then
                           objTextFile.Write(objRSEmissions.Fields("emissionValue") & ",")
                        end if
                        objRSEmissions.MoveNext
                     loop
                  else
                     Do while not objRSEmissions.EOF
                        if objRSEmissions.Fields("repNumber") = rep and _
                          objRSEmissions.Fields("parameter") = item and _
                          objRSEmissions.Fields("unitNumber") = unit and _
                          objRSEmissions.Fields("name") = plant and _
                          objRSEmissions.Fields("emission") = emission then
                           objTextFile.Write(objRSEmissions.Fields("emissionValue") & ",")
                        end if
                        objRSEmissions.MoveNext
                     loop
                  end if
                 end if
               next
               objTextFile.Write(vbcrlf)
               aReps(iClear) = ""
               iClear = iClear + 1
            end if
         next
         iClearIndex = iClearIndex + 1
       next
    end if
   Next
next
objTextFile.Close
Set objTextFile = Nothing
Set objFSO = Nothing


%>

<HTML>
<HEAD>
<TITLE>Search Results</TITLE>
<LINK href="../main.css" rel="stylesheet" type="text/css">
</HEAD>
<BODY>

<% 
   ' output a sucess message and other relevant search info for error checking
   printHeader "Export Successful"
   response.write("<strong>The following emissions were written to disk: </strong><br>")
   for each item in aEmissions
      if item <> "" then
         response.write(item & "<BR>")
      end if
   next
   response.write("<strong>The following Parameters were written to disk: </strong><BR>" & vbcrlf)
   for each item in aParameters
      if item <> "" then
         response.write(item & "<BR>")
      end if
   next
   response.write("<STRONG>The file is located at: </STRONG>" & sfilename & "<BR>" & vbcrlf)
   response.write("<A HREF=export\" & Request.Form("filename") & _
                  ">Click here to download the exported .csv file</A><BR>" & vbcrlf)
   printFooter
%>


</BODY>
</HTML>