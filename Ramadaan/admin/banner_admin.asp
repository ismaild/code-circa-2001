<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/Connections/ConnRamadaan.asp" -->
<!--#include virtual="/assets/scripts/deny.asp"-->
<% 
'	Option Explicit 
'	Response.Expires = 0
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>SiteName - Banner Administration</TITLE>
</HEAD>
<BODY>

<%
Dim objConn
Set objConn = Server.CreateObject("ADODB.Connection")
' I keep my Connection string in a App Variable, you can put your here.
objConn.ConnectionString = MM_ConnRamadaan_STRING
objConn.Open

Dim objRS
Dim strSQL
Set objRS = Server.CreateObject("ADODB.Recordset")
Set objRS = objConn.Execute("Select 1")

%>

<H3 >
Banner Administration
</H3>
<CENTER><a href="BANNER_Admin.asp?Action=List">List</a>&nbsp;&nbsp;
<a href="BANNER_Admin.asp?Action=New">New</a></CENTER><BR>
<%
' Startup Code -----------------------------------------------------------------------------------------
' Variables
Dim fAdId, fAdName, fImageURL, fAdType, fHeight, fWidth, fALTText, fHTML
Dim fWeight, fLinktoURL, fTrackImpressions, fTrackClicks, TotalWeight, WeightPercent
    
' Here is my code for posting -----------------------------------------------------------------------
If Request("Action") = "Post" Then
	' Response.Write "I'm posting " & Request("AdId")
    
	'Fix single quotes
	fHTML = Request("HTML")
	fHTML = Replace(fHTML, "'","''")
		
	' Call the stored procedure
	strSQL = "spBannerPostAd01 @pAdId=" & Request("AdId")
	strSQL = strSQL & ", @pAdName='" & Request("AdName") & "'"
	strSQL = strSQL & ", @pImageURL='" & Request("ImageURL") & "'"
	strSQL = strSQL & ", @pHeight=" & Request("Height")
	strSQL = strSQL & ", @pWidth=" & Request("Width")
	strSQL = strSQL & ", @pALTText='" & Request("ALTText") & "'"
	strSQL = strSQL & ", @pHTML='" & fHTML & "'"
	strSQL = strSQL & ", @pAdType='" & Request("AdType") & "'"
	strSQL = strSQL & ", @pWeight=" & Request("Weight")
	strSQL = strSQL & ", @pLinkToURL='" & Request("LinkToURL") & "'"
	strSQL = strSQL & ", @pTrackImpressions='" & Request("TrackImpressions") & "'"
	strSQL = strSQL & ", @pTrackClicks='" & Request("TrackClicks") & "'"
		
		
	Response.Write "<BR><BR><XMP>" & strSQL & "</XMP><BR><BR>"
		
	Set objRS = objConn.Execute(strSQL)
	Response.Write objRS("ResultSet")
		
		
		
		
	'If ok, transfer back.
	If CInt(objRS("ResultSet")) = 0 Then
		Response.Redirect("BANNER_Admin.asp?Action=List")	
	End If
    
End If
    
' Here is my code for adding / editing ---------------------------------------------------------------
If Request("Action") = "Edit" or Request("Action") = "New" Then
	' Response.Write "I'm editing " & Request("AdId")
	If Request("Action") = "Edit" Then 
		strSQL = "SELECT * from BANNER_Ads Where AdId = " & Request("AdId")
		Set objRS = objConn.Execute(strSQL)
			
		fAdId = objRS("AdId")
		fAdName = trim(objRS("AdName"))
		fImageURL = trim(objRS("ImageURL"))
		fAdType = trim(objRS("Adtype"))
		fHeight = objRS("Height")
		fWidth = objRS("Width")
		fALTText = trim(objRS("ALTText"))
		fHTML = trim(objRS("HTML"))
		fWeight = objRS("Weight")
		fLinkToURL = trim(objRS("LinkToURL"))
		fTrackImpressions = trim(objRS("TrackImpressions"))
		fTrackClicks = trim(objRS("TrackClicks"))
	Else
		fAdId = 0
		fAdType = "Image"
		fWeight = 1
		fHeight = 0
		fWidth = 0
		fTrackImpressions = "N"
		fTrackClicks = "N"
			
	End If
%>
<FORM action="BANNER_Admin.asp" id=FORM1 method=post name="BannerEdit">
<input type = "hidden" name="AdId" value="<%=fAdId%>">
<input type = "hidden" name="Action" value="Post">

<p>Ad ID: <%=fAdId%>&nbsp;&nbsp;&nbsp;&nbsp;
Name: <input type=TEXT SIZE=50 MAXLENGTH=100 Name="AdName" value="<%=fAdName%>" >
&nbsp;&nbsp;&nbsp;&nbsp;
Type <SELECT NAME=AdType>
<%
If fAdType = "Image" Then
	Response.Write "<OPTION VALUE=Image SELECTED>Image"
	Response.Write "<OPTION VALUE=Flash>Flash"
Else
	Response.Write "<OPTION VALUE=Image>Image"
	Response.Write "<OPTION VALUE=Flash SELECTED>Flash"
End If
%>
</SELECT></P>
<P>Weight: <input type=TEXT SIZE=5 Name="Weight" value="<%=fWeight%>" >&nbsp;&nbsp;&nbsp;&nbsp;
Track Impressions: <input type=TEXT SIZE=2 Name="TrackImpressions" value="<%=fTrackImpressions%>" >&nbsp;&nbsp;&nbsp;&nbsp;
Track Clicks: <input type=TEXT SIZE=2 Name="TrackClicks" value="<%=fTrackClicks%>" >
<P><B>For an Image Ad:</B>
<P>Image URL: <input type=TEXT SIZE=50 MAXLENGTH=100 Name="ImageURL" value="<%=fImageURL%>" ><BR>
LinkToURL: <input type=TEXT SIZE=50 MAXLENGTH=100 Name="LinkToURL" value="<%=fLinkToURL%>" ><BR>
ALT Text: <input type=TEXT SIZE=50 MAXLENGTH=100 Name="ALTText" value="<%=fALTText%>" ><BR>
Width: <input type=TEXT SIZE=10 Name="Width" value="<%=fWidth%>" >&nbsp;&nbsp;&nbsp;&nbsp;
Height: <input type=TEXT SIZE=10 Name="Height" value="<%=fHeight%>" ><BR>
	
<P><B>For an HTML Ad:</B><BR>
<TExtarea name="HTML" Rows=8 Cols=100><%=fHTML%></textarea></p>
	
	
<P><CENTER><input type=submit value="Update Banner" id=submit1 name=submit1></CENTER>
</FORM>
	
	
<%
		
    
End If
    
' Here is my code for Listing Ads ------------------------------------------------------------------
If Request("Action") = "List" or Request("Action") = "" Then
    
	strSQL = "SELECT TotalWeight = sum(isnull(Weight, 0)) From BANNER_Ads"
	Set objRS = objConn.Execute(strSQL)
	TotalWeight = objRS("TotalWeight")
		
    
	strSQL = "SELECT AdId, AdName, AdType, Weight, Clicks, TrackClicks "
	strSQL = strSQL & " FROM BANNER_Ads Order By Weight DESC, AdId"
	Set objRS = objConn.Execute(strSQL)
		
	%>
	<TABLE Width="100%" BORDER = 0>
	  <TR>
	    <TD>AdId</TD>
	    <TD>Ad Name</TD>
	    <TD>Type</TD>
	    <TD>Weight</TD>
	    
    <TD>Clicks</TD>
	    <TD>Track<BR>Clicks</TD>
	  </TR>
	<%
	Do Until objRS.EOF = True 
		Response.Write "<TR>"
		Response.Write "<TD>"
		Response.Write "<a href=BANNER_Admin.asp?Action=Edit&AdId=" & objRS("AdId") & ">"
		Response.Write objRS("AdId") 
		Response.Write "</a></TD>"
		Response.Write "<TD>" & objRS("AdName") & "</TD>"
		Response.Write "<TD>" & objRS("AdType") & "</TD>"
		Response.Write "<TD>" & objRS("Weight")
		If CInt(objRS("Weight")) > 0 Then
			WeightPercent = CInt(CInt(objRS("Weight")) / TotalWeight * 1000 ) / 10
			Response.Write " (" & CStr(WeightPercent) & "%)</TD>"
		End If
		Response.Write "<TD>" & objRS("Clicks") & "</TD>"
		Response.Write "<TD>" & objRS("TrackClicks") & "</TD>"
		
		Response.Write "</TR>"
		objRS.MoveNext
	Loop
	%>
	</TABLE>
	<%
End If
  
objRS.Close
Set objRS = Nothing
objConn.Close
Set objConn = Nothing  
    
%>   
        
</BODY>
</HTML>
