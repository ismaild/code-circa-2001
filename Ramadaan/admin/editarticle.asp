<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%> 
<!--#include file="../Connections/ConnRamadaan.asp" -->
<!--#include virtual="/assets/scripts/deny.asp"-->
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
' *** Update Record: set variables

If (CStr(Request("MM_update")) = "frmAdd" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = MM_ConnRamadaan_STRING
  MM_editTable = "dbo.Articles"
  MM_editColumn = "A_ID"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "/admin/articles.asp?Msg=Record Added"
  MM_fieldsStr  = "A_Title|value|A_Description|value|A_Link|value|A_Featured|value|A_Front|value"
  MM_columnsStr = "A_Title|',none,''|A_Description|',none,''|A_Link|',none,''|A_Featured|none,none,NULL|A_Front|none,none,NULL"

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
' *** Update Record: construct a sql update statement and execute it

If (CStr(Request("MM_update")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  ' create the sql update statement
  MM_editQuery = "update " & MM_editTable & " set "
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
      MM_editQuery = MM_editQuery & ","
    End If
    MM_editQuery = MM_editQuery & MM_columns(MM_i) & " = " & MM_formVal
  Next
  MM_editQuery = MM_editQuery & " where " & MM_editColumn & " = " & MM_recordId

  If (Not MM_abortEdit) Then
    ' execute the update
    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
  End If

End If
%>
<%
Dim rsArticle__MMColParam
rsArticle__MMColParam = "1"
If (Request.QueryString("AID") <> "") Then 
  rsArticle__MMColParam = Request.QueryString("AID")
End If
%>
<%
Dim rsArticle
Dim rsArticle_numRows

Set rsArticle = Server.CreateObject("ADODB.Recordset")
rsArticle.ActiveConnection = MM_ConnRamadaan_STRING
rsArticle.Source = "SELECT * FROM dbo.Articles WHERE A_ID = " + Replace(rsArticle__MMColParam, "'", "''") + ""
rsArticle.CursorType = 0
rsArticle.CursorLocation = 2
rsArticle.LockType = 1
rsArticle.Open()

rsArticle_numRows = 0
%>
<html><!-- InstanceBegin template="/Templates/admin.dwt.asp" codeOutsideHTMLIsLocked="false" -->
<head>
<!-- InstanceBeginEditable name="doctitle" -->
<title>Add Article</title>
<!-- InstanceEndEditable --> 
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../assets/scripts/admin.css" rel="stylesheet" type="text/css">
<!-- InstanceBeginEditable name="head" --><!-- InstanceEndEditable -->
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="779" border="0" cellpadding="0" cellspacing="0">
  <!--DWLayoutTable-->
  <tr> 
    <td width="779" height="24" valign="top" bgcolor="#006600" class="header"><font color="#FFFFFF">Ramadaan.co.za 
      Admin System - </font></td>
  </tr>
  <tr> 
    <td height="229" valign="top"><!-- InstanceBeginEditable name="Content" -->
	  <h4>Edit Article</h4>
      <form ACTION="<%=MM_editAction%>" METHOD="POST" name="frmAdd" id="frmAdd">
        <table width="100%" border="0">
          <tr> 
            <td width="36%"><div align="right">Title :</div></td>
            <td width="64%"><input name="A_Title" type="text" id="A_Title" value="<%=(rsArticle.Fields.Item("A_Title").Value)%>" size="40" maxlength="50"></td>
          </tr>
          <tr> 
            <td height="27"> <div align="right">Description:</div></td>
            <td><input name="A_Description" type="text" id="A_Description" value="<%=(rsArticle.Fields.Item("A_Description").Value)%>" size="50" maxlength="1000"></td>
          </tr>
          <tr> 
            <td><div align="right">Link(The file name excluding the extention)</div></td>
            <td><input name="A_Link" type="text" id="A_Link" value="<%=(rsArticle.Fields.Item("A_Link").Value)%>" size="25" maxlength="30"></td>
          </tr>
          <tr> 
            <td><div align="right">Featured Article(Please note only 1 can b featured) 
              </div></td>
            <td><select name="A_Featured" id="A_Featured">
                <option value="1">Yes</option>
                <option value="0" selected>No</option>
              </select> </td>
          </tr>
          <tr> 
            <td><div align="right">Displayed On front Page ?</div></td>
            <td><select name="A_Front" id="A_Front">
                <option value="1">Yes</option>
                <option value="0" selected>No</option>
              </select></td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
            <td><input name="Add" type="submit" id="Add" value="Add"></td>
          </tr>
        </table>
        <input type="hidden" name="MM_update" value="frmAdd">
        <input type="hidden" name="MM_recordId" value="<%= rsArticle.Fields.Item("A_ID").Value %>">
      </form>
      <!-- InstanceEndEditable --></td>
  </tr>
</table>
</body>
<!-- InstanceEnd --></html>
<%
rsArticle.Close()
Set rsArticle = Nothing
%>
