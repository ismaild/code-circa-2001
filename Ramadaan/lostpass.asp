<%@language = "VBSCRIPT" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Email Password</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="assets/scripts/ramadaan.css" rel="stylesheet" type="text/css">
<link href="/assets/scripts/menu.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/JavaScript">
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
      } else if (test!='R') { num = parseFloat(val);
        if (isNaN(val)) errors+='- '+nm+' must contain a number.\n';
        if (test.indexOf('inRange') != -1) { p=test.indexOf(':');
          min=test.substring(8,p); max=test.substring(p+1);
          if (num<min || max<num) errors+='- '+nm+' must contain a number between '+min+' and '+max+'.\n';
    } } } else if (test.charAt(0) == 'R') errors += '- '+nm+' is required.\n'; }
  } if (errors) alert('The following error(s) occurred:\n'+errors);
  document.MM_returnValue = (errors == '');
}
//-->
</script>
</head>

<body>
<% If Request.QueryString("Sent") = "" Then %>
<form action="assets/scripts/sendmail.asp" method="post" name="frm" id="frm" onSubmit="MM_validateForm('Email','','RisEmail');return document.MM_returnValue">
  <p align="center" class="maintop"><span class="text">E-mail :</span> 
    <input name="Email" type="text" class="fields" id="Email" maxlength="30">
    <input name="Submit" type="submit" class="fields" value="Email Me">
    <input name="hdnMailHandler" type="hidden" id="hdnMailHandler" value="1">
    <input name="send" type="hidden" id="send" value="True">
	<BR>
	<% 
	If Request.Querystring("MSG") <> "" Then Response.Write(Request.QueryString("MSG"))
	 %>
  </p>
</form>
<% End If%>

<% If Request.Querystring("Sent") = "True" then %>
<p align="center" class="maintop">Your Password has been sent. <a href="#" class="menu">Close this Window</a></p>
<% End If %>

</body>
</html>
