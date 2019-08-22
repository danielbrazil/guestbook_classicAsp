<%
'Dimension variables
Dim adoCon 			'Holds the Database Connection Object
Dim rsAddComments		'Holds the recordset for the new record to be added to the database
Dim strSQL			'Holds the SQL query for the database

'Create an ADO connection odject
Set adoCon = Server.CreateObject("ADODB.Connection")

'Set an active connection to the Connection object using a DSN-less connection
adoCon.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("guestbook.mdb")

'Set an active connection to the Connection object using DSN connection
'adoCon.Open "DSN=guestbook"

'Create an ADO recordset object
Set rsAddComments = Server.CreateObject("ADODB.Recordset")

'Initialise the strSQL variable with an SQL statement to query the database
strSQL = "SELECT tblComments.Name, tblComments.Comments FROM tblComments;"

'Set the cursor type we are using so we can navigate through the recordset
rsAddComments.CursorType = 2

'Set the lock type so that the record is locked by ADO when it is updated
rsAddComments.LockType = 3

'Open the tblComments table using the SQL query held in the strSQL varaiable
rsAddComments.Open strSQL, adoCon

'Tell the recordset we are adding a new record to it
rsAddComments.AddNew

'Add a new record to the recordset
rsAddComments.Fields("Name") = Request.Form("name")
rsAddComments.Fields("Comments") = Request.Form("comments")

'Write the updated recordset to the database
rsAddComments.Update

'Reset server objects
rsAddComments.Close
Set rsAddComments = Nothing
Set adoCon = Nothing

'Redirect to the guestbook.asp page
Response.Redirect "guestbook.asp"
%>