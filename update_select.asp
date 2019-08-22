<%
%>
<html>
<head>
<title>Update Entry Select</title>
</head>
<body bgcolor="white" text="black">
<%
'Dimension variables
Dim adoCon 			'Holds the Database Connection Object
Dim rsGuestbook			'Holds the recordset for the records in the database
Dim strSQL			'Holds the SQL query for the database

'Create an ADO connection odject
Set adoCon = Server.CreateObject("ADODB.Connection")

'Set an active connection to the Connection object using a DSN-less connection
adoCon.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("guestbook.mdb")

'Set an active connection to the Connection object using DSN connection
'adoCon.Open "DSN=guestbook"

'Create an ADO recordset object
Set rsGuestbook = Server.CreateObject("ADODB.Recordset")

'Initialise the strSQL variable with an SQL statement to query the database
strSQL = "SELECT tblComments.* FROM tblComments;"

'Open the recordset with the SQL query 
rsGuestbook.Open strSQL, adoCon

'Loop through the recordset
Do While not rsGuestbook.EOF
	
	'Write the HTML to display the current record in the recordset
	Response.Write ("<br>")
	Response.Write ("<a href=""update_form.asp?ID=" & rsGuestbook("ID_no") & """>")
	Response.Write (rsGuestbook("Name")) 
	Response.Write ("</a>")
	Response.Write ("<br>")
	Response.Write (rsGuestbook("Comments"))
	Response.Write ("<br>")

	'Move to the next record in the recordset
	rsGuestbook.MoveNext

Loop

'Reset server objects
rsGuestbook.Close
Set rsGuestbook = Nothing
Set adoCon = Nothing
%>
</body>
</html>