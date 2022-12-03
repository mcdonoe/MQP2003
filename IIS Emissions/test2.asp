<%
  response.write("i want to see: " & Request.Form("test"))
  response.write("<br> sorting by " & Request("sortby"))
%>