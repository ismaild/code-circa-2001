<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/Connections/ConnRamadaan.asp" -->
<% 
If Session("MM_UserName") = "" Then 
 Response.Redirect("/ecards.asp?Msg=Please Login First")
End If 
'FUNCTION RANDOM 
DIM intLow ' declare lowest number variable 
DIM intHigh ' declare highest number variable
Dim strCardNo
RANDOMIZE TIMER 
intLow = 100000000000000 ' set lowest number variable 
intHigh = 999999999999999 ' set highest number variable 
strCardNo = (FormatNumber(Int((intHigh - intLow + 1) * RND + intHigh),0,0,0,0) )
'END FUNCTION 
'RANDOM 
'Response.Write(strCardNo)
set rsCheckNo = Server.Createobject("ADODB.RecordSet")
rsCheckNo.ActiveConnection = MM_ConnRamadaan_STRING
rsCheckNo.Source = "Select C_CardNo From Cards Where C_CardNo = '" & strCardNo & "'"
rsCheckNo.Open()
If rsCheckNo.EOF And rsCheckNo.BOF Then
%>
	<%
	' *** Edit Operations: declare variables
	
	Dim MM_editAction
	Dim MM_abortEdit
	Dim MM_editQuery
	Dim MM_editCmd
	
	Dim MM_editConnection
	Dim MM_editTable
	Dim MM_editRedirectUrl
	Dim MM_editColumn
	Dim MM_recordId
	
	Dim MM_fieldsStr
	Dim MM_columnsStr
	Dim MM_fields
	Dim MM_columns
	Dim MM_typeArray
	Dim MM_formVal
	Dim MM_delim
	Dim MM_altVal
	Dim MM_emptyVal
	Dim MM_i
	
	MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
	If (Request.QueryString <> "") Then
	  MM_editAction = MM_editAction & "?" & Request.QueryString
	End If
	
	' boolean to abort record edit
	MM_abortEdit = false
	
	' query string to execute
	MM_editQuery = ""
	%>
	<%
	' *** Insert Record: set variables
	
	If (CStr(Request("MM_insert")) = "form1") Then
	
	  MM_editConnection = MM_ConnRamadaan_STRING
	  MM_editTable = "dbo.Cards"
	  MM_editRedirectUrl = "/ecards.asp?Msg=Card Sent"
	  MM_fieldsStr  = "C_To|value|C_ToEmail|value|C_From|value|C_FromEmail|value|C_URL|value|C_Message|value"
	  MM_columnsStr = "C_To|',none,''|C_ToEmail|',none,''|C_From|',none,''|C_FromEmail|',none,''|C_URL|',none,''|C_Message|',none,''"
	
	  ' create the MM_fields and MM_columns arrays
	  MM_fields = Split(MM_fieldsStr, "|")
	  MM_columns = Split(MM_columnsStr, "|")
	  
	  ' set the form values
	  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
		MM_fields(MM_i+1) = CStr(Request.Form(MM_fields(MM_i)))
	  Next
	
	  ' append the query string to the redirect URL
	  If (MM_editRedirectUrl <> "" And Request.QueryString <> "") Then
		If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
		  MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
		Else
		  MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
		End If
	  End If
	
	End If
	%>
	<%
	' *** Insert Record: construct a sql insert statement and execute it
	
	Dim MM_tableValues
	Dim MM_dbValues
	
	If (CStr(Request("MM_insert")) <> "") Then
	
	  ' create the sql insert statement
	  MM_tableValues = ""
	  MM_dbValues = ""
	  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
		MM_formVal = MM_fields(MM_i+1)
		MM_typeArray = Split(MM_columns(MM_i+1),",")
		MM_delim = MM_typeArray(0)
		If (MM_delim = "none") Then MM_delim = ""
		MM_altVal = MM_typeArray(1)
		If (MM_altVal = "none") Then MM_altVal = ""
		MM_emptyVal = MM_typeArray(2)
		If (MM_emptyVal = "none") Then MM_emptyVal = ""
		If (MM_formVal = "") Then
		  MM_formVal = MM_emptyVal
		Else
		  If (MM_altVal <> "") Then
			MM_formVal = MM_altVal
		  ElseIf (MM_delim = "'") Then  ' escape quotes
			MM_formVal = "'" & Replace(MM_formVal,"'","''") & "'"
		  Else
			MM_formVal = MM_delim + MM_formVal + MM_delim
		  End If
		End If
		If (MM_i <> LBound(MM_fields)) Then
		  MM_tableValues = MM_tableValues & ","
		  MM_dbValues = MM_dbValues & ","
		End If
		MM_tableValues = MM_tableValues & MM_columns(MM_i)
		MM_dbValues = MM_dbValues & MM_formVal
	  Next
	  MM_editQuery = "insert into " & MM_editTable & " (" & MM_tableValues & ", C_CardNo) values (" & MM_dbValues & ",'" & strCardNo & "')"
	
	  If (Not MM_abortEdit) Then
		' execute the insert
		Set MM_editCmd = Server.CreateObject("ADODB.Command")
		MM_editCmd.ActiveConnection = MM_editConnection
		MM_editCmd.CommandText = MM_editQuery
		MM_editCmd.Execute
		MM_editCmd.ActiveConnection.Close
		strNext = chr(10)
		strHost = "mail.sycon.co.za"
		strFromName = Request.Form("C_From")
		strFromEmail = "webmaster@ramadaan.co.za"
		strFromEmail2 = Request.Form("C_FromEmail")
		strTo = Request.Form("C_ToEmail") '"i.mahomed@sycon.co.za" '
		strToName = Request.Form("C_To")
		strSubject = "You have recieved an E-Card from " & strFromName 
		strBody = strToName & " You have recieved an E-Card From " & strFromName & ", To pick up your card click the link below" & strNext & strNext & "http://www.ramadaan.co.za/cardview.asp?cardno=" & strCardNo & strNext & strNext & "Or go to the URL below and enter in this Code: " & strCardNo & strNext & strNext & "http://www.ramadaan.co.za/ecards.asp" & strNext & strNext & "Send Ramadaan E-Cards from: http://www.ramadaan.co.za"
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
			strErrorURL = "/ecards.asp?Msg="
			strErr = Err.Description
			Response.Redirect(strErrorUrl & strErr)
		else
			bSuccess = True
		End If
	
		If (MM_editRedirectUrl <> "") Then
		 Response.Redirect(MM_editRedirectUrl)
		End If
	  End If
	
	End If
	%>
	<% 
	Else
	Response.Write("An Error Occured Please Try Again, Card No Exists:" & strCardNo)
	End If 
	%>
	