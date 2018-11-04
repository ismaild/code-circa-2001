<%@Language="VBSCRIPT" %>
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
      <% If Request.QueryString("ErrMsg") = "" Then %>
	  <p align="center"><b>Thank you For Your Submission</b></p>
      <p align="center"><b>Your Message Has Been Sent </b></p>
	  <% End If %>
	  <% If Request.QueryString("ErrMsg") <> "" Then %> 
      <p align="center"><b><font color="#FF0000">An Error Has Occured: <% Request.Querystring("ErrMSG") %></font></b></p>
            <p>&nbsp;</p>
	   <% End If %>
      <p></p>
      <div align="center"><a href="#" onClick="window.close()">Close Window</a></div>
    </td>
  </tr>
</table>
</body>
</html>
