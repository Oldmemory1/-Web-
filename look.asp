<html>
<body>
<%
set conn=Server.CreateObject("ADODB.Connection")
conn.Provider="Microsoft.Jet.OLEDB.4.0"
conn.Open "d:/myweb/score.mdb"
set rs = Server.CreateObject("ADODB.recordset")
rs.Open "Select * from score", conn

do until rs.EOF
    for each x in rs.Fields
    Response.Write(x.value & " ") 
    next
    Response.Write("<br />")
    rs.MoveNext
loop

rs.close
conn.close
%>
</body>
</html>
