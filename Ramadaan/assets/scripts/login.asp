<%
' *** Validate request to log in to this site.
MM_LoginAction = Request.ServerVariables("URL")
If Request.QueryString<>"" Then MM_LoginAction = MM_LoginAction + "?" + Request.QueryString
MM_valUsername=CStr(Request.Form("username"))
If MM_valUsername <> "" Then
  MM_fldUserAuthorization=""
  MM_redirectLoginSuccess= MM_LoginAction
  MM_redirectLoginFailed= MM_LoginAction 
  If Request.QueryString<>"" And Request.QueryString("Msg") = "" Then 
  MM_redirectLoginFailed= MM_LoginAction & "&MSG=Invalid Username or Password"
  End If 
  if Request.QueryString("Msg") <> "" Then 
  MM_redirectLoginFailed= MM_LoginAction '& "?MSG=Invalid Username or Password"
  MM_redirectLoginSuccess= MM_LoginAction & "&Deny=1"
  End If 
  If Request.QueryString="" Then 
  MM_redirectLoginFailed= MM_LoginAction & "?MSG=Invalid Username or Password"
  End If
  
  MM_flag="ADODB.Recordset"
  set MM_rsUser = Server.CreateObject(MM_flag)
  MM_rsUser.ActiveConnection = MM_ConnRamadaan_STRING
  MM_rsUser.Source = "SELECT U_Email, U_Password, U_FirstName"
  If MM_fldUserAuthorization <> "" Then MM_rsUser.Source = MM_rsUser.Source & "," & MM_fldUserAuthorization
  MM_rsUser.Source = MM_rsUser.Source & " FROM dbo.Users WHERE U_Email='" & Replace(MM_valUsername,"'","''") &"' AND U_Password='" & Replace(Request.Form("password"),"'","''") & "'"
  MM_rsUser.CursorType = 0
  MM_rsUser.CursorLocation = 2
  MM_rsUser.LockType = 3
  MM_rsUser.Open
  If Not MM_rsUser.EOF Or Not MM_rsUser.BOF Then 
    ' username and password match - this is a valid user
    Session("MM_Username") = MM_valUsername
	Session("IM_FirstName") = MM_rsUser("U_FirstName").Value
    If (MM_fldUserAuthorization <> "") Then
      Session("MM_UserAuthorization") = CStr(MM_rsUser.Fields.Item(MM_fldUserAuthorization).Value)
    Else
      Session("MM_UserAuthorization") = ""
    End If
    if CStr(Request.QueryString("accessdenied")) <> "" And true Then
      MM_redirectLoginSuccess = Request.QueryString("accessdenied")
    End If
    MM_rsUser.Close
    Response.Redirect(MM_redirectLoginSuccess)
  End If
  MM_rsUser.Close
  Response.Redirect(MM_redirectLoginFailed)
End If
%>
