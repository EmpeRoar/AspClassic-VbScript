<html>
	<head>
		<title>Asp Classic</title>
	</head>
	<body>
		<div>
			 <%
				Dim iNumber
				'Here is a comment
				iNumber = iNumber + 1
			%>

			Alright, Classic Asp here we go!
			<%=Time()%>
			
			
			<%  Dim iHour
				iHour = Hour(Time())
				If (iHour >= 0 And iHour < 12 ) Then %>
				Good Morning!
			<% ElseIf (iHour > 11 And iHour < 16 ) Then %>
				Good Afternoon!
			<% Else %>
				Good Evening!
			<% End If %>
		</div>	

		<% 
			Application.Contents("Name") = "Homer Simpson" 
			Session.Contents("Age") = 39


			sName = Application("Name")
			sAge = Session("Age")

			

		%>
	</body>	
</html>	