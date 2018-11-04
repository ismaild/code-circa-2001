<%@Language="VBSCRIPT" %> 
<!--#include file="Connections/ConnIMS.asp" -->
<%
set rsRefNo = Server.CreateObject("ADODB.Recordset")
rsRefNo.ActiveConnection = MM_ConnIMS_STRING
rsRefNo.Source = "SELECT * FROM REFNOS"
rsRefNo.CursorType = 0
rsRefNo.CursorLocation = 2
rsRefNo.LockType = 3
rsRefNo.Open()
rsRefNo_numRows = 0
%>
<% 
Dim RefNo
Refno = "IMS" & (rsRefNo.Fields.Item("REFNO").Value) + 1
%>
<html>
<head>
<title>Confirm Order</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="Assets/scripts/ims.css" type="text/css">
<script language="JavaScript">
<!--
function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
//-->
</script>
</head>
<body bgcolor="#FFFFFF" text="#000000" background="Assets/images/blueback.gif">
<table width="95%" border="0" cellpadding="0" cellspacing="0" height="126" align="center" class="BoxNB">
  <tr> 
    <td width="100%" height="98" valign="top" bgcolor="#FFFFFF" class="BoxNB"> 
      <p><b>Confirm Order for: <%= Request.Form("hdnApp") %></b></p>
      <p><b>Details</b></p>
      <Form name="frmSubmit" action="/Assets/scripts/aspmail.asp" method="post">
        <table width="100%" border="0" class="BoxNOB">
          <tr> 
            <td width="16%"> 
              <div align="right"><font size="2"><b>Order: </b></font></div>
            </td>
            <td width="84%"><font size="2"><%= Request.Form("hdnApp") %></font></td>
          </tr>
          <tr> 
            <td width="16%"> 
              <div align="right"><b>Order REF No:</b> </div>
            </td>
            <td width="84%"><%=(rsRefNo.Fields.Item("REFNO").Value)%> 
              <input type="hidden" name="hdnRefNo" value="<%=RefNo%>">
            </td>
          </tr>
          <tr> 
            <td width="16%"> 
              <div align="right"><font size="2"><b>Name: </b></font></div>
            </td>
            <td width="84%"> 
              <p><%= Request.Form("txtName") %> 
                <input type="hidden" name="hdnName" value="<%= Request.Form("txtName") %>">
              </p>
            </td>
          </tr>
          <tr> 
            <td height="28" width="16%"> 
              <div align="right"><font size="2"><b>Email: </b></font></div>
            </td>
            <td height="28" width="84%"><%= Request.Form("txtEmail") %> 
              <input type="hidden" name="hdnEmail" value="<%= Request.Form("txtEmail") %>">
            </td>
          </tr>
          <tr> 
            <td height="28" width="16%"> 
              <div align="right"><font size="2"><b>Cell Number: </b></font></div>
            </td>
            <td height="28" width="84%"><%= Request.Form("txtPhone") %> 
              <input type="hidden" name="hdnCell" value="<%= Request.Form("txtPhone") %>">
            </td>
          </tr>
          <tr> 
            <td height="28" width="16%"> 
              <div align="right"><b>Phone Number: </b></div>
            </td>
            <td height="28" width="84%"><%= Request.Form("txtCell") %> 
              <input type="hidden" name="hdnPhone" value="<%= Request.Form("txtCell") %>">
            </td>
          </tr>
          <tr> 
            <td height="28" width="16%"> 
              <div align="right"><b>Address: </b></div>
            </td>
            <td height="28" width="84%"><%= Request.Form("txtAddress") %> 
              <input type="hidden" name="hdnAddress" value="<%= Request.Form("txtAddress") %>">
            </td>
          </tr>
          <tr> 
            <td height="28" width="16%"> 
              <div align="right"><b> City: </b></div>
            </td>
            <td height="28" width="84%"><%= Request.Form("txtCity") %> 
              <input type="hidden" name="hdnCity" value="<%= Request.Form("txtCity") %>">
            </td>
          </tr>
          <tr> 
            <td height="28" width="16%"> 
              <div align="right"><b>Area Code: </b></div>
            </td>
            <td height="28" width="84%"><%= Request.Form("txtAreaCode") %> 
              <input type="hidden" name="hdnAreaCode" value="<%= Request.Form("txtAreaCode") %>">
            </td>
          </tr>
          <tr> 
            <td height="28" width="16%"> 
              <div align="right"><b>Comments:</b></div>
            </td>
            <td height="28" width="84%"><%= Request.Form("txtComments") %> 
              <input type="hidden" name="hdnComments" value="<%= Request.Form("txtComments") %>">
            </td>
          </tr>
          <tr> 
            <td height="28" width="16%"> 
              <div align="right"> 
                <input type="hidden" name="hdnApp" value="<%= Request.Form("hdnApp") %>">
                <input type="hidden" name="hdnTo" value="info@ftr.co.za">
                <input type="hidden" name="hdnSubject" value="Order For: <%= Request.Form("hdnApp") %>">
                <input type="hidden" name="hdnMailHandler" value="1">
                <input type="hidden" name="hdnRedirect" value="/orderrep.asp">
                <input type="hidden" name="Send" value="True">
              </div>
            </td>
            <td height="28" width="84%"> 
              <p> 
                <input type="submit" name="Cancel" value="Cancel" class="fields" onClick="MM_goToURL('parent','JavaScript:history.go(-1)');return document.MM_returnValue">
                <input type="submit" name="Submit" value="Order" class="fields">
              </p>
            </td>
          </tr>
        </table>
      </Form>
      <p>&nbsp;</p>
      <p>&nbsp;</p>
      <p></p>
      <div align="center"><a href="#" onClick="window.close()">Close Window</a></div>
    </td>
  </tr>
</table>
</body>
</html>
<%
rsRefNo.Close()
%>

