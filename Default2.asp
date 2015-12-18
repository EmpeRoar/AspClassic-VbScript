<%@ language="VBScript" %>
<%
Dim txtFirstName, txtLastName, txtEmailAddr
Dim sMessage
'**********************************************************************
'* Main
'*
'* The main subroutine for this page...
'**********************************************************************
Sub Main()
'Was this page submitted?
if ( Request("cmdSubmit") = "Submit" ) Then
'Reformat the data into a more readable format...
txtFirstName = InitCap(Request("txtFirstName"))
txtLastName = InitCap(Request("txtLastName"))
txtEmailAddr = LCase(Request("txtEmailAddr"))
'Check the email address for the correct components...
if (Instr(1, txtEmailAddr, "@") = 0 _
or Instr(1, txtEmailAddr, ".") = 0 ) Then
sMessage = "The email address you entered does not " _
& "appear to be valid."
Else
'Make sure there is something after the period..
if ( Instr(1, txtEmailAddr, ".") = Len(txtEmailAddr) _
or Instr(1, txtEmailAddr, "@") = 1 or _
(Instr(1, txtEmailAddr, ".") = Instr(1, txtEmailAddr, "@") + 1) ) Then
sMessage = "You must enter a complete email address."
end if
End If
'We passed our validation, show that all is good...
if ( sMessage = "" ) Then
sMessage = "Thank you for your input. All data has " _
& "passed verification."
else
Session("ErrorCount") = Session("ErrorCount") + 1