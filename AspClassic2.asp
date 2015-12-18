
<% 
	Application.Contents("Name") = "Homer Simpson" 
	Session.Contents("Age") = 39

	sName = Application("Name")
	sAge = Session("Age")
%>

<html>
<head>
<title>Sample ASP Page 2</title>
</head>
<body>
	<p>The Application.Contents</p>
	<%
		Dim Item
		For Each Item In Application.Contents
			Response.Write Item & " = [" & Application(Item) & "]<br>"
		Next
	%>
	
	<p>The Session.Contents</p>
	<%
		For Each Item In Session.Contents
			Response.Write Item & " = [" & Session(Item) & "]<br>"
		Next
	%>
</body>
</html>