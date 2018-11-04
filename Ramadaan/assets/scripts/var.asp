<%
Dim strRamDate, strRamStartDate, strCurrentDate, strNoDays, strYear, strMonth, strDay, strDateDiff
'strYear = Year(Now)
'strMonth = Month(Now)
'strDay = Day(Now)
'strCurrentDate cdate(strYear & "/" & strMonth & "/" & strDay)
strCurrentDate = Date()
strCurrentTime = FormatDateTime(Now, 4)
strRamStartDate = "2002/11/06"
strNoDays = 30
'If strRamStartDate = strCurrentDate Then
strRamDate = 1
'response.write strRamDate
'else
strDateDiff = DateDiff("D", cdate(strRamStartDate), cdate(strCurrentDate))
'response.write strRamStartDate & "<BR>"
'response.Write strCurrentDate & "<BR>"
'Response.Write strDateDiff & "<BR>"
strRamdate = 1 + strDateDiff
'Response.Write strRamdate
'End If
If Request.QueryString("Type") <> "" Then 
	If Request.QueryString("Type") = "1" Then 
		strHeader = "Hadith"
		strDescription = "A new hadith for every day of ramadaan. Check back here every day if you would like to submit content <a href=contact.asp class=menu>click here</a> <Br> <BR>"
		strImageDay = "/assets/images/hadithofday.gif"
	End If
	If Request.QueryString("Type") = "2" then
		strHeader = "Recipes" 
		strDescription = "A new recipe for everyday of ramadaan. Check back here every day if you would like to submit content <a href=contact.asp class=menu>click here</a> <Br> <BR>"
		strImageDay = "/assets/images/recipeofday.gif"
	End If 
End If

%>
