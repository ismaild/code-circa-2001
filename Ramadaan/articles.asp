<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="Connections/ConnRamadaan.asp" -->
<%
Dim rsArticles
Dim rsArticles_numRows

Set rsArticles = Server.CreateObject("ADODB.Recordset")
rsArticles.ActiveConnection = MM_ConnRamadaan_STRING
rsArticles.Source = "SELECT A_Title, A_Link FROM dbo.Articles ORDER BY A_ID DESC"
rsArticles.CursorType = 0
rsArticles.CursorLocation = 2
rsArticles.LockType = 1
rsArticles.Open()

rsArticles_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 10
Repeat1__index = 0
rsArticles_numRows = rsArticles_numRows + Repeat1__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

Dim rsArticles_total
Dim rsArticles_first
Dim rsArticles_last

' set the record count
rsArticles_total = rsArticles.RecordCount

' set the number of rows displayed on this page
If (rsArticles_numRows < 0) Then
  rsArticles_numRows = rsArticles_total
Elseif (rsArticles_numRows = 0) Then
  rsArticles_numRows = 1
End If

' set the first and last displayed record
rsArticles_first = 1
rsArticles_last  = rsArticles_first + rsArticles_numRows - 1

' if we have the correct record count, check the other stats
If (rsArticles_total <> -1) Then
  If (rsArticles_first > rsArticles_total) Then
    rsArticles_first = rsArticles_total
  End If
  If (rsArticles_last > rsArticles_total) Then
    rsArticles_last = rsArticles_total
  End If
  If (rsArticles_numRows > rsArticles_total) Then
    rsArticles_numRows = rsArticles_total
  End If
End If
%>
<%
Dim MM_paramName 
%>
<%
' *** Move To Record and Go To Record: declare variables

Dim MM_rs
Dim MM_rsCount
Dim MM_size
Dim MM_uniqueCol
Dim MM_offset
Dim MM_atTotal
Dim MM_paramIsDefined

Dim MM_param
Dim MM_index

Set MM_rs    = rsArticles
MM_rsCount   = rsArticles_total
MM_size      = rsArticles_numRows
MM_uniqueCol = ""
MM_paramName = ""
MM_offset = 0
MM_atTotal = false
MM_paramIsDefined = false
If (MM_paramName <> "") Then
  MM_paramIsDefined = (Request.QueryString(MM_paramName) <> "")
End If
%>
<%
' *** Move To Record: handle 'index' or 'offset' parameter

if (Not MM_paramIsDefined And MM_rsCount <> 0) then

  ' use index parameter if defined, otherwise use offset parameter
  MM_param = Request.QueryString("index")
  If (MM_param = "") Then
    MM_param = Request.QueryString("offset")
  End If
  If (MM_param <> "") Then
    MM_offset = Int(MM_param)
  End If

  ' if we have a record count, check if we are past the end of the recordset
  If (MM_rsCount <> -1) Then
    If (MM_offset >= MM_rsCount Or MM_offset = -1) Then  ' past end or move last
      If ((MM_rsCount Mod MM_size) > 0) Then         ' last page not a full repeat region
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' move the cursor to the selected record
  MM_index = 0
  While ((Not MM_rs.EOF) And (MM_index < MM_offset Or MM_offset = -1))
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend
  If (MM_rs.EOF) Then 
    MM_offset = MM_index  ' set MM_offset to the last possible record
  End If

End If
%>
<%
' *** Move To Record: if we dont know the record count, check the display range

If (MM_rsCount = -1) Then

  ' walk to the end of the display range for this page
  MM_index = MM_offset
  While (Not MM_rs.EOF And (MM_size < 0 Or MM_index < MM_offset + MM_size))
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend

  ' if we walked off the end of the recordset, set MM_rsCount and MM_size
  If (MM_rs.EOF) Then
    MM_rsCount = MM_index
    If (MM_size < 0 Or MM_size > MM_rsCount) Then
      MM_size = MM_rsCount
    End If
  End If

  ' if we walked off the end, set the offset based on page size
  If (MM_rs.EOF And Not MM_paramIsDefined) Then
    If (MM_offset > MM_rsCount - MM_size Or MM_offset = -1) Then
      If ((MM_rsCount Mod MM_size) > 0) Then
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' reset the cursor to the beginning
  If (MM_rs.CursorType > 0) Then
    MM_rs.MoveFirst
  Else
    MM_rs.Requery
  End If

  ' move the cursor to the selected record
  MM_index = 0
  While (Not MM_rs.EOF And MM_index < MM_offset)
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend
End If
%>
<%
' *** Move To Record: update recordset stats

' set the first and last displayed record
rsArticles_first = MM_offset + 1
rsArticles_last  = MM_offset + MM_size

If (MM_rsCount <> -1) Then
  If (rsArticles_first > MM_rsCount) Then
    rsArticles_first = MM_rsCount
  End If
  If (rsArticles_last > MM_rsCount) Then
    rsArticles_last = MM_rsCount
  End If
End If

' set the boolean used by hide region to check if we are on the last record
MM_atTotal = (MM_rsCount <> -1 And MM_offset + MM_size >= MM_rsCount)
%>
<%
' *** Go To Record and Move To Record: create strings for maintaining URL and Form parameters

Dim MM_keepNone
Dim MM_keepURL
Dim MM_keepForm
Dim MM_keepBoth

Dim MM_removeList
Dim MM_item
Dim MM_nextItem

' create the list of parameters which should not be maintained
MM_removeList = "&index="
If (MM_paramName <> "") Then
  MM_removeList = MM_removeList & "&" & MM_paramName & "="
End If

MM_keepURL=""
MM_keepForm=""
MM_keepBoth=""
MM_keepNone=""

' add the URL parameters to the MM_keepURL string
For Each MM_item In Request.QueryString
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepURL = MM_keepURL & MM_nextItem & Server.URLencode(Request.QueryString(MM_item))
  End If
Next

' add the Form variables to the MM_keepForm string
For Each MM_item In Request.Form
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepForm = MM_keepForm & MM_nextItem & Server.URLencode(Request.Form(MM_item))
  End If
Next

' create the Form + URL string and remove the intial '&' from each of the strings
MM_keepBoth = MM_keepURL & MM_keepForm
If (MM_keepBoth <> "") Then 
  MM_keepBoth = Right(MM_keepBoth, Len(MM_keepBoth) - 1)
End If
If (MM_keepURL <> "")  Then
  MM_keepURL  = Right(MM_keepURL, Len(MM_keepURL) - 1)
End If
If (MM_keepForm <> "") Then
  MM_keepForm = Right(MM_keepForm, Len(MM_keepForm) - 1)
End If

' a utility function used for adding additional parameters to these strings
Function MM_joinChar(firstItem)
  If (firstItem <> "") Then
    MM_joinChar = "&"
  Else
    MM_joinChar = ""
  End If
End Function
%>
<%
' *** Move To Record: set the strings for the first, last, next, and previous links

Dim MM_keepMove
Dim MM_moveParam
Dim MM_moveFirst
Dim MM_moveLast
Dim MM_moveNext
Dim MM_movePrev

Dim MM_urlStr
Dim MM_paramList
Dim MM_paramIndex
Dim MM_nextParam

MM_keepMove = MM_keepBoth
MM_moveParam = "index"

' if the page has a repeated region, remove 'offset' from the maintained parameters
If (MM_size > 1) Then
  MM_moveParam = "offset"
  If (MM_keepMove <> "") Then
    MM_paramList = Split(MM_keepMove, "&")
    MM_keepMove = ""
    For MM_paramIndex = 0 To UBound(MM_paramList)
      MM_nextParam = Left(MM_paramList(MM_paramIndex), InStr(MM_paramList(MM_paramIndex),"=") - 1)
      If (StrComp(MM_nextParam,MM_moveParam,1) <> 0) Then
        MM_keepMove = MM_keepMove & "&" & MM_paramList(MM_paramIndex)
      End If
    Next
    If (MM_keepMove <> "") Then
      MM_keepMove = Right(MM_keepMove, Len(MM_keepMove) - 1)
    End If
  End If
End If

' set the strings for the move to links
If (MM_keepMove <> "") Then 
  MM_keepMove = MM_keepMove & "&"
End If

MM_urlStr = Request.ServerVariables("URL") & "?" & MM_keepMove & MM_moveParam & "="

MM_moveFirst = MM_urlStr & "0"
MM_moveLast  = MM_urlStr & "-1"
MM_moveNext  = MM_urlStr & CStr(MM_offset + MM_size)
If (MM_offset - MM_size < 0) Then
  MM_movePrev = MM_urlStr & "0"
Else
  MM_movePrev = MM_urlStr & CStr(MM_offset - MM_size)
End If
%>
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
  <h4>Articles</h4>
  <div align="center"> 
    <% If MM_offset <> 0 Then %>
    <A HREF="<%=MM_movePrev%>"><img src="/assets/images/back.gif" alt="Go Back" width="34" height="29" border="0"></A> 
    <% End If ' end MM_offset <> 0 %>
    <A HREF="<%=MM_moveNext%>"> 
    <% If Not MM_atTotal Then %>
    <img src="/assets/images/forward.gif" alt="Go Forward" border="0"> 
    <% End If ' end Not MM_atTotal %>
    </A> 
    <% 
While ((Repeat1__numRows <> 0) AND (NOT rsArticles.EOF)) 
%>
    <% If Not rsArticles.EOF Or Not rsArticles.BOF Then %>
    <table width="100%" border="0">
      <tr> 
        <td><a href="/view.asp?Type=3&var=<%=(rsArticles.Fields.Item("A_Link").Value)%>" class="menu"><%=(rsArticles.Fields.Item("A_Title").Value)%></a></td>
      </tr>
    </table>
    <% End If ' end Not rsArticles.EOF Or NOT rsArticles.BOF %>
    <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsArticles.MoveNext()
Wend
%>
  </div>
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
rsArticles.Close()
Set rsArticles = Nothing
%>
