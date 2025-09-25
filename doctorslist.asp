#ASP.net that i used here seem to have been outdated, so I want to see how we dan rework the connections. to make the system work in MacOs environment

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
       <td><%Response.Write(x.value)%></td>
   <%	next%>
        <td><%Response.Write ("<a href=""edit.asp?ID=" & rs("studID") & """>Edit</a>")%></td>
	    <td><%Response.Write ("<a href=""del.asp?ID=" & rs("studID") & """>Delete</a>")%></td>
	    <%rs.MoveNext %>
    </tr>	
<%loop 
rs.close
conn.close
%>
<tr><td>
<form action="add.asp" method="post">
<input name="theStudName" /></td><td>
<input name="theStudID" /></td><td colspan="2">
<input type="submit" value="Add new Student" /></td></form></tr>
</table>

</body>
</html>
