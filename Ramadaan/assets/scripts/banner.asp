<%
set rsBanner = Server.CreateObject("ADODB.Recordset")
rsBanner.ActiveConnection = MM_ConnSahara_STRING
rsBanner.Source = "SELECT *  FROM BANNERS"
rsBanner.CursorType = 3
rsBanner.CursorLocation = 2
rsBanner.LockType = 1
rsBanner.Open()
rsBanner_numRows = 0

'the following codes are added to rotate the banners:
Dim rndMax
rndMax = CInt(rsBanner.RecordCount)
rsBanner.MoveFirst

Dim rndNumber
Randomize Timer
rndNumber = Int(RND * rndMax)
rsBanner.Move rndNumber
'end of codes.
%>
<% If (rsBanner.Fields.Item("B_Type").Value) = 1 Then %>
<p><a href="/assets/scripts/redirect.asp?id=<%=(rsBanner.Fields.Item("B_ID").Value)%>&url=<%=(rsBanner.Fields.Item("B_URL").Value)%>" target="_blank"><img src="<%=(rsBanner.Fields.Item("B_IMAGE").Value)%>" border="0" alt = "<%=(rsBanner.Fields.Item("B_ALT").Value)%>"></a></p>
<% End If %>
<% If (rsBanner.Fields.Item("B_Type").Value) = 2 Then %>
<p><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="468" height="60">
    <param name=movie value="<%=(rsBanner.Fields.Item("B_IMAGE").Value)%>">
    <param name=quality value=high>
    <embed src="<%=(rsBanner.Fields.Item("B_IMAGE").Value) %>" quality=high pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="468" height="60">
    </embed> 
  </object></p>
<% End If %>
<%
rsBanner.Close()
%>
