<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="Connections/ConnRamadaan.asp" -->
<!--#Include virtual="/assets/scripts/login.asp"-->
<!--#Include virtual="/assets/scripts/var.asp"-->
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
<!-- InstanceBeginEditable name="head" --><!-- InstanceEndEditable -->
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
<div id="Layer3" style="position:absolute; left:134px; top:199px; width:510px; height:100px; z-index:9; background-color: #FFFFFF; layer-background-color: #FFFFFF;" class="Content"> 
<% If Request.QueryString("Order") <> "" Then %>
	
  <h4>Order</h4>
	 <p>Please Note: This order is sent via email, You will be contacted for shipping details. Please ensure your contact details are correct.</p>
  <p align="center" class="menu">Mail: <a href="mailto:webmaster@ramadaan.co.za" class="menu">webmaster@ramadaan.co.za</a></p>
	<% else %>
  <h4>Contact Us</h4>
   <p>Feel the need to Rant at us? Got a problem with some of our content? Or would 
    you just like to compliment/ Submit content? Feel Free to contact us via Phone/Fax/Email/Telepathy</p>
  <p align="center" class="menu">Mail: <a href="mailto:webmaster@ramadaan.co.za" class="menu">webmaster@ramadaan.co.za</a></p>
  <% End IF %>
  <form action="/assets/scripts/sendmail.asp" method="post" name="frmEmail" id="frmEmail" onSubmit="MM_validateForm('txtName','','R','txtEmail','','RisEmail','txtTel','','NisNum');return document.MM_returnValue" onclick="highlight(event)">
    <table width="100%" border="0" class="text">
      <tr> 
        <td width="30%"><div align="right">Name : </div></td>
        <td width="70%"><input name="txtName" type="text" class="fields" id="txtName" size="30" maxlength="30"></td>
      </tr>
      <tr> 
        <td><div align="right">E-Mail : </div></td>
        <td><input name="txtEmail" type="text" class="fields" id="txtEmail" size="25" maxlength="30"></td>
      </tr>
      <tr> 
        <td><div align="right">Telephone :</div></td>
        <td><input name="txtTel" type="text" class="fields" id="txtTel" size="20" maxlength="10"></td>
      </tr>
      <tr> 
        <td><div align="right">
		<% If Request.QueryString("Order") = "1" Then %>
 		Choose Volume? : 		
		<% Else %>
		Query Type? : 
		<% End If %>
		</div></td>
        <td><select name="sctQuery" class="fields" id="sctQuery">
		<% If Request.QueryString("Order") = "1" Then %>
		<option value="IMP Box Set" selected>Box Set (10 CD's)</option>
		<option value="IMP Sudais and Shuraim">1 Sudais and Shuraim</option>
		<option value="IMP Sheikh Hussary">2 Sheikh Hussary </option>
		<option value="IMP Sheikh Abdul Basit">3 Sheikh Abdul Basit</option>
		<option value="IMP Ahmed Ajmi">4 Ahmed Ajmi</option>
		<option value="IMP Saad Ghamdi">5 Saad Ghamdi</option>
		<option value="IMP Sheikh Huzaifee">6 Sheikh Huzaifee</option>
		<option value="IMP Khalid Qehtani">7 Khalid Qehtani</option>
		<option value="IMP Rafaee">8 Rafaee</option>
		<option value="IMP Mohamed Hassan">9 Mohamed Hassan</option>
		<option value="IMP Aziz al-Ahmed">10 Aziz al-Ahmed</option>
		<option value="IMP Abdul Hadi Kanakar">11 Abdul Hadi Kanakar</option>
		<option value="IMP Mohamed Ayyob">12 Mohamed Ayyob</option>
		<option value="IMP Abdullah Basfar">13 Abdullah Basfar</option>
		<option value="IMP Abdul Mohsen Alqasem">14 Abdul Mohsen Alqasem</option>
		<% Else %>
            <option value="Comment" selected>Comment</option>
            <option value="Marketing">Marketing</option>
            <option value="Submit a link">Submit a link</option>
			<% End IF %>
          </select></td>
      </tr>
      <tr>
        <td><div align="right">Comments : </div></td>
        <td><textarea name="txtComments" cols="30" rows="5" class="fields" id="txtComments"><%If Request.QueryString("title") <> "" then response.Write("I Would like to Place an order for " & Request.QueryString("Title"))%></textarea></td>
      </tr>
      <tr> 
        <td>&nbsp;</td>
        <td><input name="Send" type="submit" class="fields" id="Send" value="Send">
          <input name="hdnMailHandler" type="hidden" id="hdnMailHandler" value="2"></td>
      </tr>
    </table>
  </form>
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