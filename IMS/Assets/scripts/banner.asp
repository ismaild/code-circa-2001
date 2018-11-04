<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../../Connections/ConnIMS.asp" -->
<%
set BannersRS = Server.CreateObject("ADODB.Recordset")
BannersRS.ActiveConnection = MM_ConnIMS_STRING
BannersRS.Source = "SELECT *  FROM Banners"
BannersRS.CursorType = 3
BannersRS.CursorLocation = 2
BannersRS.LockType = 3
BannersRS.Open()
BannersRS_numRows = 0
%>
<% 
Randomize Timer
BannersRS.Move Int(Rnd * Cint(BannersRS.RecordCount))
BannersRS("Shown")=BannersRS("Shown") + 1
BannersRS.Update
%>
<p><a href="/Assests/assets/scripts/RedirectMe.asp?BannerID=<%=(BannersRS.Fields.Item("BannerID").Value)%>" target="_blank"><img src="../assets/images/banners/<%=(BannersRS.Fields.Item("Image").Value)%>" border="0" alt="<%=(BannersRS.Fields.Item("Image").Value)%>"></a></p>
<%
BannersRS.Close()
%>
