<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/Connections/ConnRamadaan.asp" -->
<%

if(Request.QueryString("id") <> "") then ' spRedirect_varID = Request.QueryString("id")
	
	set spRedirect = Server.CreateObject("ADODB.Command")
	spRedirect.ActiveConnection = MM_connRamadaan_STRING
	strSql = " spBANNERUpdateAd01 @pAdId= " & (Request.QueryString("ID")) 
'	strSql = strSql & ", @pAdClicks = 2"
	spRedirect.CommandText = strSql
'	Response.Write(strSql)
	spRedirect.CommandType = 1
	spRedirect.CommandTimeout = 0
	spRedirect.Prepared = true
	spRedirect.Execute()
	
	Response.Redirect (Request.QUeryString("url")) 
End If 
%>

