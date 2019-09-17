<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8" />
    <title></title>
</head>
<body>
    <% 
        set conn=Server.CreateObject("ADODB.Connection")
        conn.Provider="Microsoft.Jet.OLEDB.4.0"
        conn.Open "c:/webdata/northwind.mdb"

        set rs = Server.CreateObject("ADODB.recordset")
        rs.Open "SELECT * FROM Customers", conn
       
        while(resultSet != null)
        {
          <p>resultSet['name']</p>
          <p>resultSet['name']</p>
        }
        rs.Close
        conn.Close
    %>
</body>
</html>