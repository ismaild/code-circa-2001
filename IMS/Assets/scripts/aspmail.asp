<%
Dim strHost
Dim strSpace
Dim strEnter
Dim strNext 
Dim strFromName
Dim strFromEmail
Dim strTo
Dim strSubject
Dim strBody
Dim strRedirectUrl 
Dim strMailHandler

Dim strEmail
Dim strCell
Dim strPhone
Dim strAddress
Dim strCity
Dim strAreaCode
Dim strComments
Dim strApp	


	strSpace = chr(32)
	strEnter = chr(13)
	strNext = chr(10)
	strHost = "127.0.0.1"

	strTo = Request.Form("hdnTo")
	strSubject = Request.Form("hdnSubject")
	strRedirectUrl = Request.Form("hdnRedirect") 

strMailHandler = Request.Form("hdnMailHandler")
If strMailHandler = "1" Then 
	strFromEmail = "webuser@ftr.co.za"
	strFromName = Request.Form("hdnName")
	strTo = Request.Form("hdnTo")
	strSubject = Request.Form("hdnSubject")
	strRedirectUrl = Request.Form("hdnRedirect") 
		
	strEmail = Request.Form("hdnEmail")
	strCell = Request.Form("hdnCell")
	strPhone = Request.Form("hdnPhone")
	strAddress = Request.Form("hdnAddress")
	strCity = Request.Form("hdnCity")
	strAreaCode = Request.Form("hdnAreaCode")
	strApp = Request.Form("hdnApp")
	strRefno = Request.Form("hdnRefNo")
	
	strBody = "Product Order" & strNext & "Name:" & strFromName & strNext & "Email:" & strEmail & strNext & "Cell Number:" & strCell & strNext & "Phone Number:" & strPhone & strNext & "Address:" & strAddress & "," & strCity & "," & strAreaCode & strNext & strNext & "Order For:" & strApp & strNext & "Thank You" 
End If 

If strMailHandler = "2"  Then 
	strTitle = Request.Form("txtTitle")
	strURL = Request.Form("txtURL")
	strDescription = Request.Form("txtDescription")
	strNPO = Request.Form("sctNPO")
	strFromName = "Web User"
	strFromEmail = "webuser@ftr.co.za"
	
	strBody = "Could You Please add this Link" & strNext & "Title: " & strTitle & strNext & "URL: " & strUrl & strNext & "Description :" & strDescription & strNext & "Non-Profit or Charity Org. ? " & strNPO & strNext & strNext & "Thank You" 
End If

If strMailHandler = "3" Then 
	strFromEmail = "webuser@ftr.co.za"
	strFromName = Request.Form("txtName")
	strTo = Request.Form("hdnTo")
	strSubject = Request.Form("hdnSubject")
	strRedirectUrl = Request.Form("hdnRedirect") 
		
	strEmail = Request.Form("txtEmail")
	strCell = Request.Form("txtCell")
	strPhone = Request.Form("txtPhone")
	strAddress = Request.Form("txtAddress")
	strCity = Request.Form("txtCity")
	strAreaCode = Request.Form("txtAreaCode")
	strComments = Request.Form("txtComments")
	
	strBody = "Comments From User" & strNext & "Name:" & strFromName & strNext & "Email:" & strEmail & strNext & "Cell Number:" & strCell & strNext & "Phone Number:" & strPhone & strNext & "Address:" & strAddress & "," & strCity & "," & strAreaCode & strNext & "Comments: " & strComments & strNext & strNext & "Thank You" 
End If

	If Request("Send") <> "" Then
		
		Set Mail = Server.CreateObject("Persits.MailSender")
		' enter valid SMTP host
		Mail.Host = strHost
		
		Mail.From = strFromEmail ' From address
		Mail.FromName = strFromName ' optional
		Mail.AddAddress strTo
		If strMailHandler = "1" Then
			Mail.AddAddress strEmail
		End If
								
		' message subject
		Mail.Subject = strSubject
		' message body
		Mail.Body = strBody
		
'Error Code
		strErr = ""
		bSuccess = False
		On Error Resume Next ' catch errors
		Mail.Send	' send message
		If Err <> 0 Then ' error occurred
			strErr = Err.Description
			strRedirectURL = "formrep.asp?ErrMsg=" & strErr
		else
			bSuccess = True
		End If
	End If
Response.Redirect(strRedirectUrl)
%>

