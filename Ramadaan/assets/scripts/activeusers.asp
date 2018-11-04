<%
'Dimension variables
Dim intDisplayDigitsLoopCount 	'Loop counter to diplay the graphical hit count


'Error handler
On Error Resume Next

'Loop to display grapical digits
For intDisplayDigitsLoopCount = 1 to Len(Application("intActiveUserNumber"))
	
	'Display the graphical active user hit count by getting the path to the image using the mid function
	Response.Write "<img src=""/assets/images/counter_images/" & Mid(Application("intActiveUserNumber"), intDisplayDigitsLoopCount, 1) & ".gif"">"
Next


'Alternative to display text output instead
'Response.Write Application("intActiveUserNumber") 

%>