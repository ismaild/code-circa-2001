<%@Language="VBSCRIPT" %> 
<% IMAppID = Request.Querystring("AID")
If IMAppID = "1" Then IM_App = "Islam Media Player" 
If IMAppID = "2" Then IM_App = "Hadith SMS" 
If IMAppID = "3" Then IM_App = "Islam Learner" 
If IMAppID = "4" Then IM_App = "Hadith Learner" 
%>
<html>
<head>
<title>Product Order</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="Assets/scripts/ims.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000" background="Assets/images/blueback.gif">
<table width="80%" border="0" cellpadding="0" cellspacing="0" height="126" align="center" class="BoxNB">
  <tr> 
    <td width="100%" height="98" valign="top" bgcolor="#FFFFFF" class="BoxNB"> 
      <p><b>Your Order will be processed soon for: <%=IM_APP%></b></p>
      <p><b>Banking Details </b></p>
      <table width="100%" border="0" class="BoxNOB">
        <tr> 
          <td width="26%"><font size="2"><b>Standard Bank </b></font></td>
          <td width="74%"><font size="2"></font></td>
        </tr>
        <tr> 
          <td width="26%"><font size="2"><b>Branch:</b></font></td>
          <td width="74%"> 
            <p><font size="2">Jacobs</font></p>
          </td>
        </tr>
        <tr> 
          <td height="28" width="26%"><font size="2"><b>Account Number: </b></font></td>
          <td height="28" width="74%"><font size="2">251794687</font></td>
        </tr>
        <tr> 
          <td height="28" width="26%"><font size="2"><b>Account Holder: </b></font></td>
          <td height="28" width="74%"><font size="2">I Vawda</font></td>
        </tr>
        <tr> 
          <td height="28" width="26%">&nbsp;</td>
          <td height="28" width="74%"> 
            <p>&nbsp;</p>
          </td>
        </tr>
      </table>
      <p>* Please note you will recieve an email with a refrence number. Please 
        use this reference number on your deposit as well as your name.</p>
      <p>&nbsp;</p>
      <p></p>
      <div align="center"><a href="#" onClick="window.close()">Close Window</a></div>
    </td>
  </tr>
</table>
</body>
</html>
