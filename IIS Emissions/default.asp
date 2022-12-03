<!--
index.asp

Overview:  This script is called upon a successful login.  It generates the
           user-specific menu given the user's ID.
Author(s): Jared F. McCaffree (jared@wpi.edu)
-->

<!--#include virtual="emissions/global.asi"-->

<%
   verifyUser
%>


<HTML>
<HEAD><TITLE>Emissions Management System: Main Page</TITLE>
</HEAD>
<LINK href="../main.css" rel="stylesheet" type="text/css">
<BODY>
<CENTER><img src="../images/covantalogo.gif" alt="Covanta Energy">
<BR>
<H1>Emissions Management System</H1>
Welcome to the Covanta Emissions Management System. 
This system keeps track of all<br>emissions testing
data from 26 plants from 1996 to the present. 
This system is to be used by authorized personnel only.</CENTER>
<HR>
<%
   Dim iUserID		' the user's ID (stored in session global variable)
   Dim iGroupID		' the user's group (stored in session global variable)
   Dim sQueryString	' the query string for the grouptable query
   Dim oConn		' connection object
   Dim objRS		' grouptable recordset object
   Dim objRSPages	' pages recordset object
   Dim sPageString	' space deliminated string of page IDs the user's group can view
   Dim asPage           ' array of pages, the result of splitting sPageString
   Dim groupName    ' the user's group name

   iUserID = Session("UserID")		' retrieve user's info
   iGroupID = Session("GroupID")
   ' select all the pages the user is allowed to view
   sQueryString = "SELECT * FROM grouptable WHERE ID = " & iGroupID

   ' connect and execute the query string
   connectDB oConn
   Set objRS = oConn.Execute(sQueryString)

   ' set the plant string, trim the whitespace and split it
   sPageString = objRs.Fields("accessiblePages")
   sPageString = trim(sPageString)
   asPage = Split(sPageString)

   'groupName for the name of the menu
   groupName = objRs.Fields("groupName")

%>
<CENTER>
<BR>
<H3>Main Menu</H3>
<B><%= objRS.Fields("groupName")%></B>
<BR>
<TABLE>
<%
				   
				   
   ' For each page the user can view query the pages table and get page's filename
   '      and link string, and print the link on the page.
   for each s in asPage
      sQueryString = "SELECT * FROM pages WHERE ID = " & s
      set objRSPages = oConn.Execute(sQueryString)
      Response.Write("<TR><TD align = center>")
      Response.Write("<A href=" & objRSPages.Fields("pageName") & ">" & _
                       objRSPages.Fields("linkText") & "</A><BR>") 
      Response.Write("</td></tr>")
   next
   disconnectDB oConn
%>
</TABLE>
<BR><HR>
<BR><IMAGE src="../images/wpimonogram.gif" alt="Developed as a WPI Major Qualifying Project">

</CENTER>
</BODY>
</HTML>