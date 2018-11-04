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
<script language="JavaScript">
<!--
function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_validateForm() { //v4.0
  var i,p,q,nm,test,num,min,max,errors='',args=MM_validateForm.arguments;
  for (i=0; i<(args.length-2); i+=3) { test=args[i+2]; val=MM_findObj(args[i]);
    if (val) { nm=val.name; if ((val=val.value)!="") {
      if (test.indexOf('isEmail')!=-1) { p=val.indexOf('@');
        if (p<1 || p==(val.length-1)) errors+='- '+nm+' must contain an e-mail address.\n';
      } else if (test!='R') {
        if (isNaN(val)) errors+='- '+nm+' must contain a number.\n';
        if (test.indexOf('inRange') != -1) { p=test.indexOf(':');
          min=test.substring(8,p); max=test.substring(p+1);
          if (val<min || max<val) errors+='- '+nm+' must contain a number between '+min+' and '+max+'.\n';
    } } } else if (test.charAt(0) == 'R') errors += '- '+nm+' is required.\n'; }
  } if (errors) alert('The following error(s) occurred:\n'+errors);
  document.MM_returnValue = (errors == '');
}
//-->
</script>
</head>

<body bgcolor="#FFFFFF" text="#000000" background="Assets/images/blueback.gif">
<table width="95%" border="0" cellpadding="0" cellspacing="0" height="100%" align="center" class="BoxNB">
  <tr> 
    <td width="100%" height="98" valign="top" bgcolor="#FFFFFF" class="BoxNB"> 
      <p><b>Place an order For: <%=IM_APP%></b></p>
      <ul>
        <li><b> Please Note Orders will only be shipped and sent once money has 
          been paid in full. </b></li>
        <li><b> Banking Details will follow once this form has been submited. 
          </b></li>
        <li><b> You will recieve a duplicate copy of your order via email </b></li>
        <li><b>Please do not include any credit card details. </b></li>
      </ul>
      <form name="form1" method="post" action="ordercon.asp" onSubmit="MM_validateForm('txtName','','R','txtEmail','','R','txtCell','','R','txtPhone','','R','txtAddress','','R','txtCity','','R','txtAreaCode','','R');return document.MM_returnValue">
        <table width="100%" border="0">
          <tr> 
            <td width="34%"> 
              <div align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Name 
                : </font></b></div>
            </td>
            <td width="66%"> 
              <input type="text" name="txtName" class="fields">
              <input type="hidden" name="hdnApp" value="<%=IM_App%>">
            </td>
          </tr>
          <tr> 
            <td width="34%"> 
              <div align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">E-Mail 
                : </font></b></div>
            </td>
            <td width="66%"> 
              <input type="text" name="txtEmail" class="fields">
            </td>
          </tr>
          <tr> 
            <td width="34%"> 
              <div align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Cell 
                Number : </font></b></div>
            </td>
            <td width="66%"> 
              <input type="text" name="txtCell" class="fields">
            </td>
          </tr>
          <tr> 
            <td width="34%"> 
              <div align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Phone 
                Number : </font></b></div>
            </td>
            <td width="66%"> 
              <input type="text" name="txtPhone" class="fields">
            </td>
          </tr>
          <tr> 
            <td width="34%"> 
              <div align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Postal 
                Address : </font></b></div>
            </td>
            <td width="66%"> 
              <input type="text" name="txtAddress" class="fields">
            </td>
          </tr>
          <tr> 
            <td width="34%"> 
              <div align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">City 
                : </font></b></div>
            </td>
            <td width="66%"> 
              <input type="text" name="txtCity" class="fields">
            </td>
          </tr>
          <tr> 
            <td width="34%"> 
              <div align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Area 
                Code : </font></b></div>
            </td>
            <td width="66%"> 
              <input type="text" name="txtAreaCode" size="6" class="fields">
            </td>
          </tr>
          <tr> 
            <td width="34%"> 
              <div align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Comments 
                : </font></b></div>
            </td>
            <td width="66%"> 
              <textarea name="txtComments" class="fields" rows="5" cols="40"></textarea>
            </td>
          </tr>
          <tr> 
            <td width="34%"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2"></font></b></td>
            <td width="66%">&nbsp;</td>
          </tr>
          <tr> 
            <td width="34%"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2"></font></b></td>
            <td width="66%"> 
              <input type="reset" name="Submit2" value="Reset" class="fields">
              <input type="submit" name="Submit" value="Submit" class="fields">
            </td>
          </tr>
        </table>
      </form>
      <p>&nbsp;</p>
      <p></p>
      <div align="center"><a href="#" onClick="window.close()">Close Window</a></div>
    </td>
  </tr>
</table>
</body>
</html>
