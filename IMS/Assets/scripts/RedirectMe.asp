<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../../Connections/ConnIMS.asp" -->
<%
Dim BannersRS__value
BannersRS__value = "1"
if (Request("BannerID")  <> "") then BannersRS__value = Request("BannerID") 
%>
<%
set BannersRS = Server.CreateObject("ADODB.Recordset")
BannersRS.ActiveConnection = MM_ConnIMS_STRING
BannersRS.Source = "SELECT *  FROM Banners  WHERE BannerID=" + Replace(BannersRS__value, "'", "''") + ""
BannersRS.CursorType = 0
BannersRS.CursorLocation = 2
BannersRS.LockType = 3
BannersRS.Open()
BannersRS_numRows = 0
%>
<%
BannersRS("Clicked")=BannersRS("Clicked")+1
BannersRS.Update()
Response.Redirect(BannersRS("URL"))
%>
<%
BannersRS.Close()
%>
