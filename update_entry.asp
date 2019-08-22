<%
'Dimension variables
Dim adoCon 			'Holds the Database Connection Object
Dim rsUpdateEntry		'Holds the recordset for the record to be updated
Dim strSQL			'Holds the SQL query for the database
Dim lngRecordNo			'Holds the record number to be updated

'Read in the record number to be updated
lngRecordNo = CLng(Request.Form("ID_no"))

'Create an ADO connection odject
Set adoCon = Server.CreateObject("ADODB.Connection")

'Set an active connection to the Connection object using a DSN-less connection
adoCon.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("guestbook.mdb")

'Set an active connection to the Connection object using DSN connection
'adoCon.Open "DSN=guestbook"

'Create an ADO recordset object
Set rsUpdateEntry = Server.CreateObject("ADODB.Recordset")

'Initialise the strSQL variable with an SQL statement to query the database
strSQL = "SELECT tblComments.* FROM tblComments WHERE ID_no=" & lngRecordNo

'Set the cursor type we are using so we can navigate through the recordset
rsUpdateEntry.CursorType = 2

'Set the lock type so that the record is locked by ADO when it is updated
rsUpdateEntry.LockType = 3

'Open the tblComments table using the SQL query held in the strSQL varaiable
rsUpdateEntry.Open strSQL, adoCon

'Update the record in the recordset
rsUpdateEntry.Fields("Name") = Request.Form("name")
rsUpdateEntry.Fields("Comments") = Request.Form("comments")

'Write the updated recordset to the database
rsUpdateEntry.Update

'Reset server objects
rsUpdateEntry.Close
Set rsUpdateEntry = Nothing
Set adoCon = Nothing

'Return to the update select page incase another record needs deleting
Response.Redirect "update_select.asp"
%>