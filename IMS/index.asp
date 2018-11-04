<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/ConnIMS.asp" -->
<%
Dim rsCharity__MMColParam
rsCharity__MMColParam = "Y"
%>
<%
set rsCharity = Server.CreateObject("ADODB.Recordset")
rsCharity.ActiveConnection = MM_ConnIMS_STRING
rsCharity.Source = "SELECT *  FROM LINKS  WHERE L_NonProfit = 'Y'  ORDER BY L_ID ASC"
rsCharity.CursorType = 0
rsCharity.CursorLocation = 2
rsCharity.LockType = 3
rsCharity.Open()
rsCharity_numRows = 0
%>
<%
set BannersRS = Server.CreateObject("ADODB.Recordset")
BannersRS.ActiveConnection = MM_ConnIMS_STRING
BannersRS.Source = "SELECT * FROM BANNERS"
BannersRS.CursorType = 3
BannersRS.CursorLocation = 2
BannersRS.LockType = 3
BannersRS.Open()
BannersRS_numRows = 0
%>
<%
set rsNews = Server.CreateObject("ADODB.Recordset")
rsNews.ActiveConnection = MM_ConnIMS_STRING
rsNews.Source = "SELECT * FROM NEWS ORDER BY N_Dated DESC"
rsNews.CursorType = 0
rsNews.CursorLocation = 2
rsNews.LockType = 3
rsNews.Open()
rsNews_numRows = 0
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = -1
Dim Repeat1__index
Repeat1__index = 0
rsCharity_numRows = rsCharity_numRows + Repeat1__numRows
%>
<%
Dim Repeat2__numRows
Repeat2__numRows = 10
Dim Repeat2__index
Repeat2__index = 0
rsNews_numRows = rsNews_numRows + Repeat2__numRows
%>
<% 
Randomize Timer
BannersRS.Move Int(Rnd * Cint(BannersRS.RecordCount))
BannersRS("Shown")=BannersRS("Shown") + 1
BannersRS.Update
%>
<html><!-- #BeginTemplate "/Templates/IMS.dwt" -->
<head>
<title>Islam Media Software</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="/Assets/scripts/ims.css" type="text/css">
<meta name="keywords" content="Islamic, Media,Software,Islam,Free,Software,Kiraat,Hadith,Quraan,Mp3,SMS,">
<meta name="Description" content="Islam Media Software Provides software solutions for the ummah  products include quraan player,hadith sms, islam learner , Free software diwbloads available">
<script language="JavaScript">
<!--
function GP_AdvOpenWindow(theURL,winName,features,popWidth,popHeight,winAlign,ignorelink,alwaysOnTop,autoCloseTime,borderless) { //v2.0
  var leftPos=0,topPos=0,autoCloseTimeoutHandle, ontopIntervalHandle, w = 480, h = 340;  
  if (popWidth > 0) features += (features.length > 0 ? ',' : '') + 'width=' + popWidth;
  if (popHeight > 0) features += (features.length > 0 ? ',' : '') + 'height=' + popHeight;
  if (winAlign && winAlign != "" && popWidth > 0 && popHeight > 0) {
    if (document.all || document.layers || document.getElementById) {w = screen.availWidth; h = screen.availHeight;}
		if (winAlign.indexOf("center") != -1) {topPos = (h-popHeight)/2;leftPos = (w-popWidth)/2;}
		if (winAlign.indexOf("bottom") != -1) topPos = h-popHeight; if (winAlign.indexOf("right") != -1) leftPos = w-popWidth; 
		if (winAlign.indexOf("left") != -1) leftPos = 0; if (winAlign.indexOf("top") != -1) topPos = 0; 						
    features += (features.length > 0 ? ',' : '') + 'top=' + topPos+',left='+leftPos;}
  if (document.all && borderless && borderless != "" && features.indexOf("fullscreen") != -1) features+=",fullscreen=1";
  if (window["popupWindow"] == null) window["popupWindow"] = new Array();
  var wp = popupWindow.length;
  popupWindow[wp] = window.open(theURL,winName,features);
  if (popupWindow[wp].opener == null) popupWindow[wp].opener = self;  
  if (document.all || document.layers || document.getElementById) {
    if (borderless && borderless != "") {popupWindow[wp].resizeTo(popWidth,popHeight); popupWindow[wp].moveTo(leftPos, topPos);}
    if (alwaysOnTop && alwaysOnTop != "") {
    	ontopIntervalHandle = popupWindow[wp].setInterval("window.focus();", 50);
    	popupWindow[wp].document.body.onload = function() {window.setInterval("window.focus();", 50);}; }
    if (autoCloseTime && autoCloseTime > 0) {
    	popupWindow[wp].document.body.onbeforeunload = function() {
  			if (autoCloseTimeoutHandle) window.clearInterval(autoCloseTimeoutHandle);
    		window.onbeforeunload = null;	}  
   		autoCloseTimeoutHandle = window.setTimeout("popupWindow["+wp+"].close()", autoCloseTime * 1000); }
  	window.onbeforeunload = function() {for (var i=0;i<popupWindow.length;i++) popupWindow[i].close();}; }   
  document.MM_returnValue = (ignorelink && ignorelink != "") ? false : true;
}
//-->
</script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bordercolor="#666666">
  <tr> 
    <td width="748" height="93" valign="top"><img src="/Assets/images/topbanner.gif" width="748" height="93"></td>
    <td width="100%" valign="top" background="/Assets/images/blueback.gif">&nbsp;</td>
  </tr>
  <tr> 
    <td height="0"></td>
    <td></td>
  </tr>
  <tr> 
    <td height="1"><img height="1" width="748" src="/Assets/spacer.gif"></td>
    <td></td>
  </tr>
</table>
<table width="100%" border="0" cellpadding="0" cellspacing="0" mm:layoutgroup="true">
  <tr> 
    <td width="180" valign="top" height="300" class="Box"> 
      <table width="85%" border="1" align="center" bgcolor="#FFFFFF" bordercolor="#CCCCCC">
        <tr> 
          <td bordercolor="#660066" bgcolor="#CCCCCC"> 
            <div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b> 
              <font color="#660066">Menu</font></b> </font></div>
          </td>
        </tr>
        <tr> 
          <td bordercolor="#FFFFFF"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b>::</b></font> 
            <font size="2" face="Verdana, Arial, Helvetica, sans-serif"><a href="/index.asp">Home 
            - Latest News</a></font></td>
        </tr>
        <tr> 
          <td bordercolor="#FFFFFF"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b>::</b> 
            <a href="/about.asp">About Us</a></font></td>
        </tr>
        <tr> 
          <td bordercolor="#FFFFFF"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b>::</b> 
            <a href="/freesoft.asp">Free software</a></font></td>
        </tr>
        <tr> 
          <td bordercolor="#FFFFFF"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b>::</b> 
            <a href="/contact.asp">Contact Us</a></font></td>
        </tr>
        <tr> 
          <td bordercolor="#FFFFFF"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>::</b> 
            <a href="/links.asp">Links</a></font></td>
        </tr>
        <tr> 
          <td bordercolor="#660066" bgcolor="#CCCCCC"> 
            <div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b><font color="#660066">Products</font></b></font></div>
          </td>
        </tr>
        <tr> 
          <td bordercolor="#FFFFFF"><font face="Verdana, Arial, Helvetica, sans-serif"><font size="2"><b>::</b> 
            <a href="/imp.asp">Islam Media Player</a></font></font></td>
        </tr>
        <tr> 
          <td bordercolor="#FFFFFF"><font face="Verdana, Arial, Helvetica, sans-serif"><font size="2"><b>::</b> 
            <a href="/hadithlearner.asp">Hadith Learner </a></font></font></td>
        </tr>
        <tr> 
          <td bordercolor="#FFFFFF"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b>::</b> 
            <a href="/islamlearner.asp">Islam Learner</a></font></td>
        </tr>
        <tr> 
          <td bordercolor="#FFFFFF"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b>::</b> 
            <a href="/adhaansoftware.asp">Adhaan Software</a></font></td>
        </tr>
        <tr> 
          <td bordercolor="#FFFFFF"> 
            <div align="left"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b>::</b> 
              <a href="/hadithsms.asp">Hadith SMS</a></font></div>
          </td>
        </tr>
      </table>
      <div align="center"></div>
    </td>
    <td width="100%" rowspan="2" valign="top" class="Content"><!-- #BeginEditable "Content" --> 
      <p><b><font size="2">Welcome To Islam Media Software On the Web</font></b></p>
      <% 
While ((Repeat2__numRows <> 0) AND (NOT rsNews.EOF)) 
%>
      <p><font size="3"><font size="2"><%=(rsNews.Fields.Item("N_Title").Value)%></font></font><b><BR>
        </b><font size="1">[ <%=(rsNews.Fields.Item("N_Dated").Value)%> - <a href="<%=(rsNews.Fields.Item("N_URL").Value)%>" target="_blank">Read 
        More</a> ]</font></p>
      <p>- <%=(rsNews.Fields.Item("N_Description").Value)%> - <a href="<%=(rsNews.Fields.Item("N_URL").Value)%>" target="_blank">Read 
        More</a> </p>
      <% 
  Repeat2__index=Repeat2__index+1
  Repeat2__numRows=Repeat2__numRows-1
  rsNews.MoveNext()
Wend
%>
      <!-- #EndEditable --></td>
    <td width="150" valign="top" class="Box"> 
      <div align="center"> 
        <p><a href="/Assets/scripts/RedirectMe.asp/<%=(BannersRS.Fields.Item("BannerID").Value)%>" target="_blank"><img border="0" alt="<%=(BannersRS.Fields.Item("Image").Value)%>" src="/Assets/images/Banners/<%=(BannersRS.Fields.Item("Image").Value)%>"></a></p>
        <p><a href="http://www.ftr.co.za" target="_blank"><img src="/Assets/images/88x31ftr.gif" width="88" height="31" border="0"></a></p>
        <p>&nbsp;</p>
      </div>
    </td>
  </tr>
  <tr> 
    <td valign="top" class="LinksList" rowspan="2"> 
      <div align="center"> 
        <p><b>Charity Organisations</b></p>
      </div>
      <% 
While ((Repeat1__numRows <> 0) AND (NOT rsCharity.EOF)) 
%>
      <p align="center"><a href="<%=(rsCharity.Fields.Item("L_URL").Value)%>" target="_blank"> 
        |<%=(rsCharity.Fields.Item("L_LinkTitle").Value)%> | </a></p>
      <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsCharity.MoveNext()
Wend
%>
      <p align="center"><a href="#" onClick="GP_AdvOpenWindow('../linkadd.htm','IMP','fullscreen=no,toolbar=no,location=no,status=no,menubar=no,scrollbars=yes,resizable=no',500,400,'center','ignoreLink','',0,'');return document.MM_returnValue">Add 
        a Link</a></p>
    </td>
    <td height="114"></td>
  </tr>
  <tr> 
    <td height="12"></td>
    <td></td>
  </tr>
  <tr> 
    <td height="31"></td>
    <td></td>
    <td></td>
  </tr>
  <tr> 
    <td height="1"><img height="1" width="180" src="/Assets/spacer.gif"></td>
    <td></td>
    <td><img height="1" width="150" src="/Assets/spacer.gif"></td>
  </tr>
</table>
</body>
<!-- #EndTemplate --></html>
<%
rsCharity.Close()
%>
<%
BannersRS.Close()
%>
<%
rsNews.Close()
%>
