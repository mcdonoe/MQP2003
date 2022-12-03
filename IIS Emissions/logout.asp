<!--
logout.asp

Overview: This file removes all session information about the user and displays
          a logout message.
Author(s): Jared F. McCaffree
-->

<% Session.Abandon %>

<HTML>
<HEAD>
<TITLE>Advanced Search for Emissions</TITLE>
<LINK href="../main.css" rel="stylesheet" type="text/css">

<BODY>

  <H3> You Have Successfully Logged Out</H3>

<HR>

<BR>
Please close the browser window for added security.
<BR><BR><BR>
<H3>
<A HREF="login.asp">Log back in</A>
</H3>

</BODY>
</HTML>