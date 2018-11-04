<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file = "connections/ConnRamadaan.asp" -->
<% 
Set rsGetVariables = Server.CreateObject("ADODB.RecordSet")
MySQL = "Select * From Messages where m_id = 2"
rsGetVariables.Activeconnection = MM_ConnRamadaan_STRING
rsGetVariables.Source = MySQL 
rsGetVariables.Open()
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body> 
From = <%= rsGetVariables("M_From").value %>
To  = <%= rsGetVariables("M_To").value %>
Message = <%= rsGetVariables("M_Message").value %>
</body>
</html>
