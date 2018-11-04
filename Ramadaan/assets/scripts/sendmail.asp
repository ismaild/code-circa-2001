<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/Connections/ConnRamadaan.asp" -->
<% 
strHost = "mail.sycon.co.za"
strSpace = chr(32)
strEnter = chr(13)
strNext = chr(10)

If Request.Form("hdnMailHandler") = "1" Then 
	Dim MySQL, StrUsername, strFirstname, strPassword
	Set rsGetUser = Server.CreateObject("ADODB.RecordSet")
	MySQL = "Select U_Email, U_Password, U_FirstName From Users where U_Email='"
	MySQL = MySQL & Request.Form("Email") & "'"
	rsGetUser.Activeconnection = MM_ConnRamadaan_STRING
	rsGetUser.Source = MySQL
	rsGetUser.Open()
	if rsGetUser.EOF or rsGetUser.BOF Then 
		strSendPass = 0 
		Response.Redirect("/lostpass.asp?MSG=That user does not exist.")
		rsGetUser.Close()
	Else 
		strSendpass = 1
		strFromEmail = "password@ramadaan.co.za"
		strFromName = "Ramadaan.co.za"
		strSubject = "Your Ramadaan.co.za password"
		strTo = rsGetUser("U_Email").value
		
		strFirstName = rsGetUser("U_Firstname").value
		strPassword = rsGetUser("U_Password").value
		
		strBody = strFirstName & strSpace & "your Ramadaan.co.za details are" & strNext & strNext & "Username(Email) : " & strTo & strNext  & "Password : " & strPassword & strNext & strNext & "http://www.ramadaan.co.za" 
		strRedirectUrl = "/lostpass.asp?Sent=True"
		strErrorUrl = "/lostpass.asp?Msg="
		rsGetUser.Open()
	End If 
End If
If Request.Form("hdnMailHandler") = "2" Then 
	Dim strName, strEmail, strQuery, strComment, strTel
	strSendpass = 2
	strName = Request.Form("txtName")
	strEmail = Request.Form("txtEmail")
	strTel = Request.Form("txtTel")
	strQuery = Request.Form("sctQuery")
	strComments = Request.Form("txtComments")

	strFromEmail = strEmail
	strFromName = strName
	strSubject = strQuery & " Query From Ramadaan.co.za"
	strTo = "webmaster@ramadaan.co.za"
	
	strBody = "The following user has submited a form on http://www.ramadaan.co.za" & strNext & strNext & "Name : " & strName & strNext & "E-Mail : " & strEmail & strNext & "Telephone : " & strTel & strNext & "Query Type : " & strQuery & strNext & strNext & "Comments : " & strComments & strNext & strNext & "Thank You"
	strRedirectUrl = "/contact.asp?Msg=Mail Sent"
	strErrorUrl = "/contact.asp?Msg=" 
	
End If

If Request("Send") <> "" & strSendPass <> 0 Then
		
		Set Mail = Server.CreateObject("Persits.MailSender")
		' enter valid SMTP host
		Mail.Host = strHost

		Mail.From = strFromEmail ' From address
		Mail.FromName = strFromName ' optional
		Mail.AddAddress strTo
				
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
			Response.Redirect(strErrorUrl & strErr)
'			response.Write(strErr)
'			strRedirectURL = "/lostpass.asp?MSG=" & strErr
		else
			bSuccess = True
			Response.Redirect(strRedirectUrl)
		End If
	End If

%>
