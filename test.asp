<html>
<body>
        <%
        set conn=Server.CreateObject("ADODB.Connection")
        conn.Provider="Microsoft.Jet.OLEDB.4.0"
        conn.Open "d:/myweb/score.mdb"
        
        sql="INSERT INTO score (ID,"
        sql=sql & "s_name)"
        sql=sql & " VALUES "
        sql=sql & "('" & Request.Form("score") & "',"
        sql=sql & "'" & Request.Form("username") & "')"
        
        on error resume next
        conn.Execute sql,recaffected
        if err<>0 then
          Response.Write("No update permissions!")
        else
          Response.Write("<h3>" & recaffected & " record added</h3>")
        end if
        conn.close
        %>
</body>               
</html>