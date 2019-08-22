<%
'Dimension variables
Dim adoCon 			'Holds the Database Connection Object
Dim rsGuestbook			'Holds the recordset for the record to be updated
Dim strSQL			'Holds the SQL query for the database
Dim lngRecordNo			'Holds the record number to be updated

'Read in the record number to be updated
lngRecordNo = CLng(Request.QueryString("ID"))

'Create an ADO connection odject
Set adoCon = Server.CreateObject("ADODB.Connection")

'Set an active connection to the Connection object using a DSN-less connection
adoCon.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("guestbook.mdb")

'Set an active connection to the Connection object using DSN connection
'adoCon.Open "DSN=guestbook"

'Create an ADO recordset object
Set rsGuestbook = Server.CreateObject("ADODB.Recordset")

'Initialise the strSQL variable with an SQL statement to query the database
strSQL = "SELECT tblComments.* FROM tblComments WHERE ID_no=" & lngRecordNo

'Open the recordset with the SQL query 
rsGuestbook.Open strSQL, adoCon
%>
<html>
<head>
<title>Guestbook Update Form</title>
</head>
<body bgcolor="white" text="black">
<!-- Begin form code -->
<form name="form" method="post" action="update_entry.asp">
  Name: <input type="text" name="name" maxlength="20" value="<% = rsGuestbook("Name") %>">  
  <br>
  Comments: <input type="text" name="comments" maxlength="60" value="<% = rsGuestbook("Comments") %>">
  <input type="hidden" name="ID_no" value="<% = rsGuestbook("ID_no") %>">
  <input type="submit" name="Submit" value="Submit">
</form>
<!-- End form code -->
</body>
</html>
<%
'Reset server objects
rsGuestbook.Close
Set rsGuestbook = Nothing
Set adoCon = Nothing
%>
