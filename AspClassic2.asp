
<% Response.Buffer = True %>
<% 
	Application.Contents("Name") = "Homer Simpson" 
	Session.Contents("Age") = 39

	sName = Application("Name")
	sAge = Session("Age")
%>

<html>
<head>
<title>Sample ASP Page 2</title>

<script language="VBScript" runat="Server">
Sub Application_OnStart
	'Globals...
		Application("ErrorPage") = "handleError.asp"
		Application("SiteBanAttemptLimit") = 10
		Application("AccessErrorPage") = "handleError.asp"
		Application("RestrictAccess") = False
	'Keep track of visitors…
	Application("NumVisits") = Application("NumVisits") + 1
End Sub
</script>

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

	<%
		'Removing Items from Content Collection
		Application("MySign") = "Gemini"
		Application.Contents.Remove("MySign")
	%>
	<%
		'Remove all
		Application.Contents.RemoveAll
	%>
	<%
		ClientIPAddress = Request("REMOTE_ADDR")
	%>

	<%=Request.QueryString("buyer")%> 

	or 

	<% 
		Response.Write(Request.QueryString("bookname"))
	%>


	<form action="book.asp" method="POST">
		Type your name: <input type="TEXT" name="buyer"><br>
		Type your requested book: <input type="TEXT" name="bookname" size="40"><br>
		<input type="SUBMIT" value="Submit">
	</form>
	<%=Request.Form("buyer")%>


	<%=Request.ServerVariables("HTTP_USER_AGENT")%>

	<%
		For Each key in Request.ServerVariables
		Response.Write "<b>" & (Key) &"</b>&nbsp"
		Response.Write (Request.ServerVariables(key)) & "<br>"
		Next
	%>

	<% Response.Cookies("BookBought") = "VBScript Programmers Reference" %>
	<%
		Response.Cookies("BookBought")("1") = "VBScript Programmers Reference"
		Response.Cookies("BookBought")("2") = "XSLT Programmers Reference"
	%>

	<% Response.Cookies("BookBought").Expires = #31-Dec-04# %>
</body>
</html>