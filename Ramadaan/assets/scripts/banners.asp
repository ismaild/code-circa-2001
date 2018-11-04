<%
'Sub BannerDisplay (MM_ConnRamadaan_STRING)

Dim bobjConn
Set bobjConn = Server.CreateObject("ADODB.Connection")
bobjConn.ConnectionString = MM_ConnRamadaan_STRING
bobjConn.Open


Dim bobjRS
Dim bstrSQL
Set bobjRS = Server.CreateObject("ADODB.Recordset")

Set bobjRS = bobjConn.Execute("spBANNERDisplay01")


If Trim(bobjRS("AdType")) = "Flash" Then
' 	Response.Write "flash Banner" 'bobjRS("HTML")
	Response.Write("<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0' width='" & trim(bobjRS("Width")) & "' height='" & trim(bobjRS("Height")) & "'>")
	Response.Write("<param name='movie' value='" & trim(bobjRS("ImageURL")) & "'>")
	Response.Write("<param name='quality' value='high'>")
	Response.Write("<embed src='" & trim(bobjRS("ImageURL")) & "' quality='high' pluginspage='http://www.macromedia.com/go/getflashplayer' type='application/x-shockwave-flash' width='" & trim(bobjRS("Width")) & "' height='" & trim(bobjRS("Height")) & "'>")
	Response.Write("</embed></object>")
End If

If Trim(bobjRS("AdType")) = "Image" Then
	Response.Write "<a href=/assets/scripts/redirect.asp?ID=" & trim(bobjRS("AdID")) & "&Url=" & trim(bobjRS("LinkToURL")) & ">"
	Response.Write "<IMG SRC=" & trim(bobjRS("ImageURL")) & " BORDER=0 "
	Response.Write " WIDTH=" & trim(bobjRS("Width"))
	Response.Write " HEIGHT=" & trim(bobjRS("Height")) 
	Response.Write " ALT=" & Chr(34) & trim(bobjRS("ALTText")) & Chr(34)
	Response.Write "></a>"
End If



bobjRS.Close
Set bobjRS = Nothing
bobjConn.Close
Set bobjConn = Nothing

'End Sub
%>