<%
 Class clsDatabase

'declare the variables 

Dim ConnectionString
Dim conn
Dim rs
Dim cmd
Dim SQL
Dim strsql
Function connect()
	'define the connection string, specify database driver
		ConnectionString = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=localhost; DATABASE=data1; " &_
		"UID=root;PASSWORD=; OPTION=3"

	'create an instance of the ADO connection and recordset objects
		Set conn = Server.CreateObject("ADODB.Connection")
		Set rs = Server.CreateObject("ADODB.Recordset")
		Set cmd = Server.CreateObject("ADODB.Command")
	'Open the connection to the database
		conn.Open ConnectionString

End Function

Function queryDatabase()
	Dim rs
'declare the SQL statement that will query the database
	strsql = "SELECT * from repairs"
'Open the recordset object executing the SQL statement and return records 
	rs.Open strsql,conn	
End Function





Function displayRecords()
'first of all determine whether there are any records 
If rs.EOF Then 
Response.Write("No records returned.") 
Else 
'if there are records then loop through the fields 
Response.write("<table border=1>")
Do While NOT rs.Eof   
Response.write("<tr>")
Response.write "<td>" +rs("RecordStatus")+"</td>"
Response.write "<td>" +rs("Name1")+"</td>"
Response.write "<td>" +rs("Name2")+ "</td>" 
Response.write("</tr>") 
rs.MoveNext     
Loop
Response.write("</table>")
End If

End Function

Function close()
	'close the connection and recordset objects freeing up resources
	'rs.Close : Set rs = Nothing : conn.Close : Set conn = Nothing
	Set rs = Nothing
	'rs.Close
	Set conn = Nothing
	'conn.Close
End Function
End Class
%>