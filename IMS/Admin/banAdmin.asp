<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="../Connections/ConnIMS.asp" -->
<%
set rsBanners = Server.CreateObject("ADODB.Recordset")
rsBanners.ActiveConnection = MM_ConnIMS_STRING
rsBanners.Source = "SELECT * FROM BANNERS"
rsBanners.CursorType = 0
rsBanners.CursorLocation = 2
rsBanners.LockType = 3
rsBanners.Open()
rsBanners_numRows = 0
%>
<%
Function MM_getParams(KeepURL,KeepForm,Params)
  RemoveList = Array("index","move")
  For index = 0 to UBound(Params) Step 2
    If (Params(index+1) <> "") Then
      If (MM_getParams <> "") Then MM_getParams = MM_getParams & "&"
      MM_getParams = MM_getParams & Params(index) & "=" & Params(index+1)
    End If
  Next
  If (KeepURL) Then
    For Each Item In Request.QueryString
      Found = False
      For Each Elem In RemoveList
        If (CStr(Elem) = CStr(Item)) Then Found = True
      Next
      If (Not Found) Then
        For index = 0 to UBound(Params) Step 2
          If (CStr(Params(index)) = CStr(Item)) Then Found = True
        Next
      End If
      If (Not Found) Then
        If (MM_getParams <> "") Then MM_getParams = MM_getParams & "&"
        MM_getParams = MM_getParams & Item & "=" & Request.QueryString(Item)
      End If
    Next
  End If
  If (KeepForm) Then
    For Each Item In Request.Form
      Found = False
      For Each Elem In RemoveList
        If (CStr(Elem) = CStr(Item)) Then Found = True
      Next
      If (Not Found) Then
        For index = 0 to UBound(Params) Step 2
          If (CStr(Params(index)) = CStr(Item)) Then Found = True
        Next
      End If
      If (Not Found) Then
        If (MM_getParams <> "") Then MM_getParams = MM_getParams & "&"
        MM_getParams = MM_getParams & Item & "=" & Request.Form(Item)
      End If
    Next
  End If
End Function
%>
<html>
<head>
<title>Banner Admin</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#FFFFFF">
<div id="Layer2" style="position:absolute; width:250px; height:42px; z-index:2; left: 146px; top: 24px"><a href="banInsert.asp">Add 
  new banner</a></div>
<div id="Layer1" style="position:absolute; width:520px; height:115px; z-index:1; left: 129px; top: 129px"> 
  <table width="94%" border="1">
    <tr> 
      <td>URL</td>
      <td>Clicked</td>
      <td>Shown</td>
    </tr>
    <%while not rsBanners.EOF%>
    <tr> 
      <td><A HREF="banUpdate.asp?<%=MM_getParams(true,false,Array("id",rsBanners.Fields.Item("BannerID").Value))%>"><%=(rsBanners.Fields.Item("URL").Value)%></A></td>
      <td><%=(rsBanners.Fields.Item("Clicked").Value)%></td>
      <td><%=(rsBanners.Fields.Item("Shown").Value)%></td>
    </tr>
    <%rsBanners.MoveNext
wend
%>
  </table>
</div>
</body>
</html>
<%
rsBanners.Close()
%>
