<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="Connections/ConnRamadaan.asp" -->
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

If (CStr(Request("MM_update")) = "frmregister" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = MM_ConnRamadaan_STRING
  MM_editTable = "dbo.Users"
  MM_editColumn = "U_ID"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "/index.asp?MSG=Details Updated"
  MM_fieldsStr  = "u_email|value|u_password2|value|U_FirstName|value|U_Lastname|value|U_Address|value|U_Address2|value|U_City|value|select|value|U_Code|value|U_Tel|value|U_TelCell|value|U_Recieve|value"
  MM_columnsStr = "U_Email|',none,''|U_Password|',none,''|U_FirstName|',none,''|U_LastName|',none,''|U_Address|',none,''|U_Address2|',none,''|U_City|',none,''|U_Prov|none,none,NULL|U_Code|',none,''|U_Tel|',none,''|U_TelCell|',none,''|U_Receive|none,none,NULL"

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
<!--#include virtual="/assets/scripts/login.asp" -->
<!--#Include virtual="/assets/scripts/var.asp"-->
<%
Dim rsUpdate__MMColParam
rsUpdate__MMColParam = "1"
If (Session("MM_Username") <> "") Then 
  rsUpdate__MMColParam = Session("MM_Username")
  Else 
  Response.Redirect("error.asp?Msg=Please login first")
End If
%>
<%
Dim rsUpdate
Dim rsUpdate_numRows

Set rsUpdate = Server.CreateObject("ADODB.Recordset")
rsUpdate.ActiveConnection = MM_ConnRamadaan_STRING
rsUpdate.Source = "SELECT * FROM dbo.Users WHERE U_Email = '" + Replace(rsUpdate__MMColParam, "'", "''") + "'"
rsUpdate.CursorType = 0
rsUpdate.CursorLocation = 2
rsUpdate.LockType = 1
rsUpdate.Open()

rsUpdate_numRows = 0
%>
<html><!-- InstanceBegin template="/Templates/ramadaan.dwt.asp" codeOutsideHTMLIsLocked="false" -->
<head>
<!-- InstanceBeginEditable name="doctitle" -->
<title>Ramadaan.co.za - </title>
<!-- InstanceEndEditable --> 
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);

var highlightcolor="#ffffff"

var ns6=document.getElementById&&!document.all
var previous=''
var eventobj
var intended=/INPUT|TEXTAREA|SELECT|OPTION/
function checkel(which){
if (which.style&&intended.test(which.tagName)){
if (ns6&&eventobj.nodeType==3)
eventobj=eventobj.parentNode.parentNode
return true
}
else
return false
}
function highlight(e){
eventobj=ns6? e.target : event.srcElement
if (previous!=''){
if (checkel(previous))
previous.style.backgroundColor=''
previous=eventobj
if (checkel(eventobj))
eventobj.style.backgroundColor=highlightcolor
}
else{
if (checkel(eventobj))
eventobj.style.backgroundColor=highlightcolor
previous=eventobj
}
}

function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}

function alertmsg() { 
alert("<%=Request.QueryString("MSG")%>");
}

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
<link href="assets/scripts/ramadaan.css" rel="stylesheet" type="text/css">
<link href="assets/scripts/menu.css" rel="stylesheet" type="text/css">
<meta content="For muslim users about during the month of ramadaan">
<meta description="For muslim users about during the month of ramadaan">
<meta keywords="Ramadaan, Ramadhaan, Ramadan, fasting, muslim, islamic, adhaan, software, eid, fitr, namaaz, prayer">
<!-- InstanceBeginEditable name="head" -->
<script language="JavaScript">
<!--
var highlightcolor="#ffffff"

var ns6=document.getElementById&&!document.all
var previous=''
var eventobj
var intended=/INPUT|TEXTAREA|SELECT|OPTION/
function checkel(which){
if (which.style&&intended.test(which.tagName)){
if (ns6&&eventobj.nodeType==3)
eventobj=eventobj.parentNode.parentNode
return true
}
else
return false
}
function highlight(e){
eventobj=ns6? e.target : event.srcElement
if (previous!=''){
if (checkel(previous))
previous.style.backgroundColor=''
previous=eventobj
if (checkel(eventobj))
eventobj.style.backgroundColor=highlightcolor
}
else{
if (checkel(eventobj))
eventobj.style.backgroundColor=highlightcolor
previous=eventobj
}
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_validateForm() { //v4.0
  var i,p,q,nm,test,num,min,max,errors='',args=MM_validateForm.arguments;
   if (document.frmregister.u_password.value != document.frmregister.u_password2.value)
  	{
	errors+='- ' + 'Your passwords do not match\n';
    }
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
<!-- InstanceEndEditable -->
</head>
<body bgcolor="#D5DBBF" text="#333333" link="#003300" vlink="#999999" alink="#FF0000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" <% If Request.QueryString("Msg") <> "" And Request.QueryString("Deny") = "" Then response.Write("onLoad=alertmsg();")%>>
<div id="Layer1" style="position:absolute; left:2px; top:119px; width:120px; height:20px; z-index:4" class="TopLeft">:: 
  Menu </div>
<div id="Layer2" style="position:absolute; left:2px; top:139px; width:120px; height:53px; z-index:5" class="SideLeft"> 
  <table width="100%" border="0">
    <tr> 
      <td><a href="index.asp" class="menu"> Home</a></td>
    </tr>
    <tr> 
      <td><a href="articles.asp" class="menu"> Articles</a></td>
    </tr>
    <tr> 
      <td><a href="/list.asp?Type=1" class="menu">Hadith</a></td>
    </tr>
    <tr> 
      <td><a href="/list.asp?Type=2" class="menu"> Recipes</a></td>
    </tr>
    <tr> 
      <td><a href="/ecards.asp" class="menu"> E-Cards</a></td>
    </tr>
    <tr>
      <td><a href="/view.asp?Type=4&var=eidsms" class="menu">Eid SMS</a></td>
    </tr>
    <tr> 
      <td><a href="/view.asp?Type=4&var=zakah" class="menu"> Zakah</a></td>
    </tr>
    <tr> 
      <td><a href="/view.asp?Type=4&var=fitr" class="menu">Fitr</a></td>
    </tr>
    <tr> 
      <td class="menu"><a href="downloads.asp" class="menu">Downloads</a></td>
    </tr>
    <tr> 
      <td class="menu"><a href="/software.asp" class="menu">Software</a></td>
    </tr>
    <tr> 
      <td class="menu"><a href="/view.asp?Type=4&var=links" class="menu">Links</a></td>
    </tr>
    <tr> 
      <td><a href="/contact.asp" class="menu">Contact Us</a></td>
    </tr>
    <tr> 
      <td><a href="/view.asp?Type=4&var=aboutus" class="menu">About Us</a></td>
    </tr>
  </table>
</div>
<div id="Layer4" style="position:absolute; left:656px; top:120px; width:120; height:20px; z-index:6" class="TopLeft">:: 
  Sponsors </div>
<div id="Layer5" style="position:absolute; left:656px; top:138px; width:120px; height:50px; z-index:7" class="SiderRight">
  <p><a href="http://www.sycon.co.za" target="_blank"><img src="assets/images/120x60_sycon.gif" width="118" height="58" border="0"></a><BR>
 <a href="http://www.ftr.co.za" target="_blank"><img src="/assets/images/ftrlogo.gif" border="0"></a> 
  </p>
</div>
<div id="Layer6" style="position:absolute; left:134px; top:120px; width:510px; height:42px; z-index:8" class="maintop"> 
  <Form ACTION="<%=MM_LoginAction%>" method="POST" name="frmLogin" id="frmLogin" onSubmit="MM_validateForm('username','','RisEmail','password','','R');return document.MM_returnValue" onclick="highlight(event)">
    <table width="100%" border="0" class="text">
      <tr> 
        <td width="53%" height="40"><h3>Welcome <% If Session("IM_FirstName") <> ""  Then Response.Write(Session("IM_FirstName"))%> ...</h3></td>
        <% If Session("MM_Username") = "" Then %>
		<td width="22%">Email: 
          <input name="username" type="text" class="fields" id="username" size="15" maxlength="30"> 
          <br> </td>
        <td width="18%">Password: 
          <input name="password" type="password" class="fields" id="password" size="10"></td>
        <td width="7%"><br> <input name="Submit" type="submit" class="fields" value="Go"></td>
		<% End If %>
		<% If Session("MM_Username") <> "" Then %>	  
        <td width="18%"><a href="assets/scripts/logout.asp" class="menu">Logout</a></td>
		<td width="18%"><a href="/update.asp" class="menu">Edit 
          Profile</a></td>
		</tr>
   		 </table>
		  <div align="center"><font color="#FFFFCC">.</font></div>
		<% End If %>
      </tr>
    </table>
	<% If Session("MM_Username") = "" Then %>
	<div align="center">
    Not a member? <a href="register.asp" class="menu">Register 
      Now!</a> -  Forgot Your Password? <a href="javascript:;" class="menu" onClick="MM_openBrWindow('/lostpass.asp','LostPass','scrollbars=yes,width=400,height=100')">Get 
      it here! </a></div>
	  <% End If %>
  </Form>
</div>
<div id="Layer7" style="position:absolute; left:656px; top:269px; width:120px; height:20px; z-index:10" class="TopLeft">:: 
  Info </div>
<div id="Layer8" style="position:absolute; left:656px; top:289px; width:120px; height:30px; z-index:11" class="SiderRight"> 
  <div align="center">
    <p>[ Active Users ]<BR>
	<!--#include virtual="/assets/scripts/activeusers.asp" --> 
	<BR>
    [ Best Viewed ]<br>
      IE 4+, 800x600 Res</p>
	  <!-- Absolute Statistics Code Start (OID=10848)-->
<script><!--
var dt = new Date();
document.write('<a target="_blank" href="http://stats.absol.co.za/showstats.asp?ownerid=10848"><img src="http://stats.absol.co.za/image.asp?ownerid=10848&pid='+escape(location.href)+'&atc='+Math.random()+'&ref='+escape(document.referrer)+'&ttl='+escape(document.title)+'&tzn='+dt.getTimezoneOffset()+'&js=yes" border=0></a>');
//-->
</script><NOSCRIPT><a target="_blank" href="http://stats.absol.co.za/showstats.asp?ownerid=10848"><IMG SRC="HTTP://stats.absol.co.za/image.asp?ownernum=10848&js=no"></a></NOSCRIPT>
<!-- Absolute Statistics Code End -->
  
  </div>
</div>
<!-- InstanceBeginEditable name="Content" --> 
<div id="Layer3" style="position:absolute; left:134px; top:199px; width:510px; height:100px; z-index:9;" class="Content"> 
  <h5>Update My Profile</h5>
  <p>Please note: To be eligible for prizes valid details need to be entered here.</p>
  <Form ACTION="<%=MM_editAction%>" METHOD="POST" name="frmregister" id="frmRegister" onSubmit="MM_validateForm('u_email','','RisEmail','u_password','','R','u_password2','','R','U_FirstName','','R','U_Lastname','','R','U_Address','','R','U_City','','R','U_Code','','RisNum','U_Tel','','RisNum');return document.MM_returnValue" onclick="highlight(event)">
    <table width="100%" border="0" class="text">
      <tr> 
        <td width="33%"><div align="right">E-mail :</div></td>
        <td width="67%"><input name="u_email" type="text" class="fields" id="u_email" value="<%=(rsUpdate.Fields.Item("U_Email").Value)%>" size="30" maxlength="50"></td>
      </tr>
      <tr> 
        <td height="21"><div align="right">Password :</div></td>
        <td><input name="u_password" type="text" class="fields" id="u_password" value="<%=(rsUpdate.Fields.Item("U_Password").Value)%>" size="20" maxlength="20"></td>
      </tr>
      <tr> 
        <td><div align="right">Password : </div></td>
        <td><input name="u_password2" type="text" class="fields" id="u_password2" value="<%=(rsUpdate.Fields.Item("U_Password").Value)%>" size="20" maxlength="20"></td>
      </tr>
      <tr> 
        <td><div align="right">First Name :</div></td>
        <td><input name="U_FirstName" type="text" class="fields" id="U_FirstName" value="<%=(rsUpdate.Fields.Item("U_FirstName").Value)%>" size="30" maxlength="30"></td>
      </tr>
      <tr> 
        <td><div align="right">Last Name :</div></td>
        <td><input name="U_Lastname" type="text" class="fields" id="U_Lastname" value="<%=(rsUpdate.Fields.Item("U_LastName").Value)%>" size="30" maxlength="30"></td>
      </tr>
      <tr> 
        <td><div align="right">Address :</div></td>
        <td><input name="U_Address" type="text" class="fields" id="U_Address" value="<%=(rsUpdate.Fields.Item("U_Address").Value)%>" size="50" maxlength="50"></td>
      </tr>
      <tr> 
        <td><div align="right">Address :</div></td>
        <td><input name="U_Address2" type="text" class="fields" id="U_Address2" value="<%=(rsUpdate.Fields.Item("U_Address2").Value)%>" size="50" maxlength="50"></td>
      </tr>
      <tr> 
        <td><div align="right">City :</div></td>
        <td><input name="U_City" type="text" class="fields" id="U_City" value="<%=(rsUpdate.Fields.Item("U_City").Value)%>" size="30" maxlength="30"></td>
      </tr>
      <tr> 
        <td><div align="right">Province :</div></td>
        <td><select name="select" class="fields">
            <option value="1">Non SA Province</option>
            <option value="2" selected>Gauteng</option>
            <option value="3">Kwa-Zulu Natal</option>
            <option value="4">Mpumalanga</option>
            <option value="5">North west</option>
            <option value="6">Eastern Cape</option>
            <option value="7">Free State</option>
            <option value="8">Nothern Province</option>
            <option value="9">Western Cape</option>
            <option value="10">Nothern Cape</option>
          </select></td>
      </tr>
      <tr> 
        <td height="17"> <div align="right">Area Code :</div></td>
        <td><input name="U_Code" type="text" class="fields" id="U_Code" value="<%=(rsUpdate.Fields.Item("U_Code").Value)%>" size="6" maxlength="4"></td>
      </tr>
      <tr> 
        <td><div align="right">Telephone No. :</div></td>
        <td><input name="U_Tel" type="text" class="fields" id="U_Tel" value="<%=(rsUpdate.Fields.Item("U_Tel").Value)%>" size="15" maxlength="10"></td>
      </tr>
      <tr> 
        <td><div align="right">Cellular No. : </div></td>
        <td><input name="U_TelCell" type="text" class="fields" id="U_TelCell" value="<%=(rsUpdate.Fields.Item("U_TelCell").Value)%>" size="15" maxlength="10"> 
        </td>
      </tr>
      <tr> 
        <td><div align="right">Recieve Special offers? : </div></td>
        <td>Yes 
          <input type="radio" name="U_Recieve" value="1" class="fields" checked>
          No 
          <input type="radio" name="U_Recieve" value="0" class="fields"></td>
      </tr>
      <tr> 
        <td>&nbsp;</td>
        <td> <input name="Update" type="submit" class="fields" id="Update" value="Update"></td>
      </tr>
    </table>
    <input type="hidden" name="MM_update" value="frmregister">
    <input type="hidden" name="MM_recordId" value="<%= rsUpdate.Fields.Item("U_ID").Value %>">
  </Form>
  <p>&nbsp;</p>
</div>
<!-- InstanceEndEditable -->
<table border="0" cellpadding="0" cellspacing="0" width="779">
  <!-- fwtable fwsrc="main.png" fwbase="main.gif" fwstyle="Dreamweaver" fwdocid = "742308039" fwnested="0" -->
  <tr> 
    <td><img src="assets/images/spacer.gif" width="78" height="1" border="0" alt=""></td>
    <td><img src="assets/images/spacer.gif" width="227" height="1" border="0" alt=""></td>
    <td><img src="assets/images/spacer.gif" width="468" height="1" border="0" alt=""></td>
    <td><img src="assets/images/spacer.gif" width="6" height="1" border="0" alt=""></td>
    <td><img src="assets/images/spacer.gif" width="1" height="1" border="0" alt=""></td>
  </tr>
  <tr> 
    <td colspan="4"><img name="main_r1_c1" src="assets/images/main_r1_c1.gif" width="779" height="5" border="0" alt=""></td>
    <td><img src="assets/images/spacer.gif" width="1" height="5" border="0" alt=""></td>
  </tr>
  <tr> 
    <td rowspan="2" colspan="2"><a href="http://www.ramadaan.co.za"><img name="main_r2_c1" src="assets/images/main_r2_c1.gif" width="305" height="65" border="0" alt=""></a></td>
    <td><!--#include virtual="/assets/scripts/banners.asp"--></td>
    <td rowspan="2"><img name="main_r2_c4" src="assets/images/main_r2_c4.gif" width="6" height="65" border="0" alt=""></td>
    <td><img src="assets/images/spacer.gif" width="1" height="60" border="0" alt=""></td>
  </tr>
  <tr> 
    <td><img name="main_r3_c3" src="assets/images/main_r3_c3.gif" width="468" height="5" border="0" alt=""></td>
    <td><img src="assets/images/spacer.gif" width="1" height="5" border="0" alt=""></td>
  </tr>
  <tr> 
    <td><img name="main_r4_c1" src="assets/images/main_r4_c1.gif" width="78" height="22" border="0" alt=""></td>
    <td colspan="3" valign="top" bgcolor="#ffffff"><table border="0" cellpadding="0" cellspacing="0" width="700">
        <!-- fwtable fwsrc="timesbann.png" fwbase="timesban.gif" fwstyle="Dreamweaver" fwdocid = "742308039" fwnested="0" -->
        <tr> 
          <td><img src="assets/images/spacer.gif" alt="" name="undefined_2" width="25" height="1" border="0"></td>
          <td><img src="assets/images/spacer.gif" alt="" name="undefined_2" width="158" height="1" border="0"></td>
          <td><img src="assets/images/spacer.gif" alt="" name="undefined_2" width="35" height="1" border="0"></td>
          <td><img src="assets/images/spacer.gif" alt="" name="undefined_2" width="40" height="1" border="0"></td>
          <td><img src="assets/images/spacer.gif" alt="" name="undefined_2" width="35" height="1" border="0"></td>
          <td><img src="assets/images/spacer.gif" alt="" name="undefined_2" width="93" height="1" border="0"></td>
          <td><img src="assets/images/spacer.gif" alt="" name="undefined_2" width="35" height="1" border="0"></td>
          <td><img src="assets/images/spacer.gif" alt="" name="undefined_2" width="41" height="1" border="0"></td>
          <td><img src="assets/images/spacer.gif" alt="" name="undefined_2" width="35" height="1" border="0"></td>
          <td><img src="assets/images/spacer.gif" alt="" name="undefined_2" width="88" height="1" border="0"></td>
          <td><img src="assets/images/spacer.gif" alt="" name="undefined_2" width="35" height="1" border="0"></td>
          <td><img src="assets/images/spacer.gif" alt="" name="undefined_2" width="41" height="1" border="0"></td>
          <td><img src="assets/images/spacer.gif" alt="" name="undefined_2" width="35" height="1" border="0"></td>
          <td><img src="assets/images/spacer.gif" alt="" name="undefined_2" width="4" height="1" border="0"></td>
          <td><img src="assets/images/spacer.gif" alt="" name="undefined_2" width="1" height="1" border="0"></td>
        </tr>
        <tr> 
          <td><img name="timesbann_RDate" src="/assets/images/times/timesbann_RDate<%=strRamDate%>.gif" width="25" height="20" border="0" alt=""></td>
          <td><img name="timesban_r1_c2" src="assets/images/timesban_r1_c2.gif" width="158" height="20" border="0" alt=""></td>
          <td><img name="timesbann_jhb_s" src="/assets/images/times/timesbann_jhb_s<%=strRamDate%>.gif" width="35" height="20" border="0" alt=""></td>
          <td><img name="timesban_r1_c4" src="assets/images/timesban_r1_c4.gif" width="40" height="20" border="0" alt=""></td>
          <td><img name="timesbann_jhb_i" src="/assets/images/times/timesbann_jhb_i<%=strRamDate%>.gif" width="35" height="20" border="0" alt=""></td>
          <td><img name="timesban_r1_c6" src="assets/images/timesban_r1_c6.gif" width="93" height="20" border="0" alt=""></td>
          <td><img name="timesbann_dbn_s" src="/assets/images/times/timesbann_dbn_s<%=strRamDate%>.gif" width="35" height="20" border="0" alt=""></td>
          <td><img name="timesban_r1_c8" src="assets/images/timesban_r1_c8.gif" width="41" height="20" border="0" alt=""></td>
          <td><img name="timesbann_dbn_i" src="/assets/images/times/timesbann_dbn_i<%=strRamDate%>.gif" width="35" height="20" border="0" alt=""></td>
          <td><img name="timesban_r1_c10" src="assets/images/timesban_r1_c10.gif" width="88" height="20" border="0" alt=""></td>
          <td><img name="timesbann_cpt_s" src="/assets/images/times/timesbann_cpt_s<%=strRamDate%>.gif" width="35" height="20" border="0" alt=""></td>
          <td><img name="timesban_r1_c12" src="assets/images/timesban_r1_c12.gif" width="41" height="20" border="0" alt=""></td>
          <td><img name="timesbann_cpt_i" src="/assets/images/times/timesbann_cpt_i<%=strRamDate%>.gif" width="35" height="20" border="0" alt=""></td>
          <td><img name="timesban_r1_c14" src="assets/images/timesban_r1_c14.gif" width="4" height="20" border="0" alt=""></td>
          <td><img src="assets/images/spacer.gif" alt="" name="undefined_2" width="1" height="20" border="0"></td>
        </tr>
      </table>
      
    </td>
    <td><img src="assets/images/spacer.gif" width="1" height="22" border="0" alt=""></td>
  </tr>
</table>
<p class="text">&nbsp;</p>
</body>
<!-- InstanceEnd --></html>
<%
rsUpdate.Close()
Set rsUpdate = Nothing
%>
