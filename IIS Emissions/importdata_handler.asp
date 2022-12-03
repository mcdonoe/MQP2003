<!--
importdata_handler.asp

Overview: This file enables the System Operator to select an csv file
          and import it into the emissions database.
Author(s): Jared F. McCaffree
           Ronald Cormier
-->

<!--#include virtual="emissions/global.asi"-->


<HTML>
<HEAD>
<%
   'verifyUser
%>

<TITLE>Import Data</TITLE>
<LINK href="../main.css" rel="stylesheet" type="text/css">

<BODY>

<%
   Dim objFSO        ' file system object
   Dim objTextFile   ' file object
   Dim aImport       ' array buffer of one line of data read from file
   Dim bFirst        ' boolean to skip the first line of the file
   Dim sInString     ' input buffer
   Dim oConn         ' connection object
   Dim sQueryString  ' query string
   Dim iPlantID      ' target plant ID
   Dim sPath         ' path to import file
   Dim sFileName     ' filename of import file
   Dim objRS         ' recordset object
   Dim sDate	     ' date of the import
   Dim sEmissions    ' string of emissions that were measured
   Dim aEmissions    ' array of emissions that were measured
   Dim bStart        ' boolean to determine when to start paying attention
   Dim iEmCounter    ' keep track of number of emissions measured
   Dim sFormula      ' formula string
   Dim sParameter    ' parameter string
   Dim iUnit         ' unit number int
   Dim iRep          ' rep  number int
   Dim sEmVals       ' string of actual emission values
   Dim aEmVals       ' array of actual emission values
   Dim iBaseCell     ' column that is next to column contains the first
                     ' emission value
   
   sDate = Date()


   bFirst = 1
   ' save the filename
   sFileName = sPathToUpload & Request.Form("fileName")

   ' Create an instance of the the File System Object and assign it to objFSO.
   Set objFSO = CreateObject("Scripting.FileSystemObject")
   ' Open the file 
   Set objTextFile = objFSO.OpenTextFile(sFileName)

   ' connect to the SQL database
   connectDB oConn

   ' Select all of the plant names and get the ID of the target plant.
   '   The plant ID is then set in the emissions database for each imported
   '   emission.
   sQueryString = "SELECT * FROM plant WHERE plantName='" & Request.Form("plant") & "'"
   set objRS = oConn.Execute(sQueryString)
   do while not objRS.EOF
      iPlantID = objRS.Fields("ID")
      objRS.movenext
   loop

   ' get the emissions that were measured, check them and put in an array
   iEmCounter = 0


   ' find cell w/ "Results" and get rest of cells starting recording from there
   sInString = objTextFile.Readline     ' read the first line
   aImport = Split(sInString, Chr(34))    ' remove quotes
   sInString = Join(aImport, "")
   aImport = Split(sInString, ",")
   bStart = 0
   iBaseCell = 0
   for each cell in aImport
      if (Instr(cell, "Results") <> 0) then
         bStart = 1    ' next cell will contain name of emission being measured
      end if
      if (bStart = 1) then    ' this contains the name of an emission measured
         if (cell <> "") then
            sEmissions = sEmissions & cell & ","
         end if
      else
         iBaseCell = iBaseCell + 1
      end if
      'Response.Write("*" & cell & " *")
   next
   'Response.Write("base cell: " & iBaseCell & "<BR>")
   sEmissions = Trim(sEmissions)
   ' remove last extra comma
   sEmissions = Left(sEmissions, Len(sEmissions) - 1)
   aEmissions = Split(sEmissions, ",")
   iEmCounter = 0
   for each emission in aEmissions
      'Response.Write(iEmCounter & "&nbsp;" & emission & "&nbsp;&nbsp;&nbsp;" & vbcrlf)
      iEmCounter = iEmCounter + 1
   next
   'Response.Write("<BR>")


   ' get the formula used
   ' find cell w/ "Parameter" and get next cell
   sInString = objTextFile.Readline     ' get the second line
   aImport = Split(sInString, ",")
   bStart = 0
   for each cell in aImport
      if (bStart = 1) then
         if (cell <> "")  then
            sFormula = cell
            bStart = 0
         end if
      end if
      if (Instr(cell, "Parameter") <> 0) then
         bStart = 1     ' next cell will contain the parameter
      end if
   next
   'Response.Write("Formula: " & sFormula & "<BR>")

   Dim iTmpBase
   Do While Not objTextFile.AtEndofStream

      ' get the parameter
      sInString = objTextFile.Readline      ' get the third line
      aImport = Split(sInString, ",")
      Dim iImportCnt   'number of cells there are in a row
      iImportCnt = 0
      for each cell in aImport
         iImportCnt = iImportCnt + 1
      next
      if (iImportCnt > 3) then
         if (aImport(3) <> "") then       'parameter must be in col 4
            sParameter = aImport(3)
         end if
      end if

      ' get the unit number
      if (iImportCnt > 1) then
         if (aImport(1) <> "") then       'unit number must be in col 2
            iUnit = aImport(1)
         end if
      end if

      ' get the rep number
      if (iImportCnt > 2) then            'rep number must be in col 3
         iRep = aImport(2)
      end if

      ' get the rest of the emission values
      ' build a string of them then split into array
      Dim iCnt, sTmpVal, aTmpVal
      iTmpBase = iBaseCell
      iCnt = 0
      sEmVals = ""
      do while (iCnt < iEmCounter)
         if (iImportCnt > iTmpBase) then
            'interpret scientific notation
            sTmpVal = aImport(iTmpBase)
            if (InStr(sTmpVal, "E") <> 0) then
               aTmpVal = Split(sTmpVal, "E")
               sTmpVal = trim(aTmpVal(0)) * (10^trim(aTmpVal(1)))
            end if
            if (sTmpVal <> "") then
               sTmpVal = CDbl(sTmpVal)
            end if
            'Response.write("here" & sTmpVal & "<BR>" & vbcrlf)
            sEmVals = sEmVals & sTmpVal & ","
         end if
         iCnt = iCnt + 1
         iTmpBase = iTmpBase + 1
      loop
      if (len(sEmVals) > 1) then
         sEmVals = Left(sEmVals, Len(sEmVals)-1)
      end if
      aEmVals = Split(sEmVals, ",")


      if (iRep <> "" AND sInString <> "") then
         'Response.Write("Parameter: " & sParameter & "&nbsp;&nbsp;&nbsp;")
         'Response.Write("Unit Num: " & iUnit & "&nbsp;&nbsp;&nbsp;")
         'Response.Write("Rep Num: " & iRep & "&nbsp;&nbsp;&nbsp;")
         'Response.Write("<BR>" & vbcrlf)

         iCnt = 0
         for each emission in aEmissions 
            sQueryString = "INSERT INTO emission(emission,emissionValue,unitNumber,repNumber," _
                         & "testMethod,plantID,importDate, parameter) values('" & emission & "','" _
                         & aEmVals(iCnt) & "'," & iUnit & "," & iRep & ",'" & sFormula & "'," _
                         & iPlantID & ",'" & sDate & "','" & sParameter & "')"
            'Response.Write(sQueryString & "<BR>" & vbcrlf)
            oConn.Execute(sQueryString)
            iCnt = iCnt + 1
         next
         'Response.Write("<BR>")
      end if
   Loop


   ' Close the file.
   objTextFile.Close

   ' Release reference to the text file.
   Set objTextFile = Nothing

   ' Release reference to the File System Object.
   Set objFSO = Nothing
%>

<H3>Import Successful!</H3><HR>
<a href="./">Back to Menu</a>
</BODY>
</HTML>