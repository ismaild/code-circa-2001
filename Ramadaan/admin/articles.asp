<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../Connections/ConnRamadaan.asp" -->
<!--#include virtual="/assets/scripts/deny.asp"-->
<%
Dim rsArticles
Dim rsArticles_numRows

Set rsArticles = Server.CreateObject("ADODB.Recordset")
rsArticles.ActiveConnection = MM_ConnRamadaan_STRING
rsArticles.Source = "SELECT A_ID, A_Title FROM dbo.Articles"
rsArticles.CursorType = 0
rsArticles.CursorLocation = 2
rsArticles.LockType = 1
rsArticles.Open()

rsArticles_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
rsArticles_numRows = rsArticles_numRows + Repeat1__numRows
%>
<html><!-- InstanceBegin template="/Templates/admin.dwt.asp" codeOutsideHTMLIsLocked="false" -->
<head>
<!-- InstanceBeginEditable name="doctitle" -->
<title>Articles</title>
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
      <p>View Articles - <a href="/admin/addarticle.asp">Add Article</a></p>
      <table width="100%" border="1" bordercolor="#FFFFFF">
        <tr> 
          <td colspan="2" bgcolor="#003300"><font color="#FFFFFF">Title</font></td>
        </tr>
        <% If Not rsArticles.EOF Or Not rsArticles.BOF Then %>
        <% 
While ((Repeat1__numRows <> 0) AND (NOT rsArticles.EOF)) 
%>
        <tr bordercolor="#000000"> 
          <td width="58%"><%=(rsArticles.Fields.Item("A_Title").Value)%></td>
          <td width="9%"><a href="/admin/editarticle.asp?AID=<%=(rsArticles.Fields.Item("A_ID").Value)%>">Edit</a></td>
        </tr>
        <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsArticles.MoveNext()
Wend
%>
        <% End If ' end Not rsArticles.EOF Or NOT rsArticles.BOF %>
      </table>
      <p><font color="#FF0000"><% If Request.QueryString("Msg") <> "" Then Response.Write(Request.QueryString("Msg"))%></font></p>
      <!-- InstanceEndEditable --></td>
  </tr>
</table>
</body>
<!-- InstanceEnd --></html>
<%
rsArticles.Close()
Set rsArticles = Nothing
%>
