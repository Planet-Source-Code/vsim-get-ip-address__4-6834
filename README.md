<div align="center">

## Get IP ADDRESS


</div>

### Description

This is usefull program to track people's IP.

Visit Count for each visiting IP address. Assume that the table IPAddresses(IPAddress, Count) exists in the DSN yourDSN.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[vsim](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/vsim.md)
**Level**          |Intermediate
**User Rating**    |4.4 (22 globes from 5 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[ASP Server Object Model](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/asp-server-object-model__4-32.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/vsim-get-ip-address__4-6834/archive/master.zip)





### Source Code

```
<html>
<head>
<title>Visit Counts per ip addresses</title>
</head>
<body>
<%
'  Get Remote IP Address.
IPAddress = Request.ServerVariables("Remote_Addr")
'  Open database connection.
Set conn = Server.CreateObject("adodb.connection")
conn.open "yourDsn"
'  Check to see the IP address exits.
'  Inefficient: to be modified.
sql = "select * from IPAddresses" & _
   " where IPAddress = '" & IPAddress & "'"
set result = conn.execute(sql)12:09 AM 8/8/01
if (not result.eof) then ' one result
  Count = Clng(result.fields.item("Count")) + 1
  sql = "update IPAddresses set Count = " & Count & _
     " where IPAddress = '" & IPAddress & "'"
else
  Count = 1
  sql = "insert into IPAddresses values('" & _
     IPAddress & "', " & Count & ")"
end if
conn.execute(sql)
conn.close
'  Print result.
Response.write("IP Address " & IPAddress & " has visited " & _
"this page for " & count & " times.")
%>
</body>
</html>
```

