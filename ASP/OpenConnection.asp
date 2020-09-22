<%
dim db, rs
set db = Server.CreateObject("ADODB.Connection")
db.open("DRIVER={Microsoft Access Driver (*.mdb)};DBQ=" & Server.MapPath("Settings.mdb") & ";")
%>