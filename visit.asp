<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Untitled Document</title>
</head>

<body>

<%
set conn=Server.CreateObject("ADODB.Connection")
conn.Provider="Microsoft.Jet.OLEDB.4.0"
conn.Open(Server.Mappath("db/project.accdb"))
set rs = Server.CreateObject("ADODB.recordset")
sql="SELECT ID, Firstname  FROM student"
rs.Open sql, conn
%>

	

<table border="1" width="100%" bgcolor="#fff5ee">
<tr>
<%for each x in rs.Fields
    response.write("<th align='left' bgcolor='#b0c4de'>" & x.name & "</th>")
next%>
<th></th><th></th></tr>
<%do until rs.EOF%>
    <tr>
    <%for each x in rs.Fields%>
<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
Dim conn, rs, sql

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Server.MapPath("db/project.accdb") & ";Persist Security Info=False;"

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT ID, Firstname FROM student"
rs.Open sql, conn
%>

<html>
<head><title>Students</title></head>
<body>
<table border="1" width="100%" bgcolor="#fff5ee">
<tr>
<% For Each x In rs.Fields %>
    <th align="left" bgcolor="#b0c4de"><%= x.Name %></th>
<% Next %>
<th></th><th></th>
</tr>

<% Do Until rs.EOF %>
<tr>
    <% For Each x In rs.Fields %>
        <td><%= x.Value %></td>
    <% Next %>
    <td><a href="edit.asp?ID=<%= rs("ID") %>">Edit</a></td>
    <td><a href="del.asp?ID=<%= rs("ID") %>">Delete</a></td>
</tr>
<% 
    rs.MoveNext 
Loop
rs.Close
conn.Close
%>

<tr>
<form action="add.asp" method="post">
    <td><input name="theStudName" /></td>
    <td><input name="theStudID" /></td>
    <td colspan="2"><input type="submit" value="Add new Student" /></td>
</form>
</tr>
</table>
</body>
</html>

