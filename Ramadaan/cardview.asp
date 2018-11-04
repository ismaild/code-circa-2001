<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="Connections/ConnRamadaan.asp" -->
<%
Dim rsCard__MMColParam
rsCard__MMColParam = "1"
If (Request.QueryString("CardNo") <> "") Then 
  rsCard__MMColParam = Request.QueryString("CardNo")

Dim rsCard
Dim rsCard_numRows

Set rsCard = Server.CreateObject("ADODB.Recordset")
rsCard.ActiveConnection = MM_ConnRamadaan_STRING
rsCard.Source = "SELECT * FROM dbo.Cards WHERE C_CardNo = '" + Replace(rsCard__MMColParam, "'", "''") + "'"
rsCard.CursorType = 0
rsCard.CursorLocation = 2
rsCard.LockType = 1
rsCard.Open()
rsCard_numRows = 0
	if Not rsCard.EOF or Not rsCard.BOF Then 
	strTo = (rsCard.Fields.Item("C_To").Value)
	strToEmail = (rsCard.Fields.Item("C_ToEmail").Value)
	strFrom = (rsCard.Fields.Item("C_From").Value)
	strFromEmail = (rsCard.Fields.Item("C_FromEmail").Value)
	strMessage = (rsCard.Fields.Item("C_Message").Value)
	strUrl = (rsCard.Fields.Item("C_Url").Value)
	End If 
	If rsCard.EOF or rsCard.BOF Then 
	Response.Redirect("/ecards.asp?Msg=Card Does not Exist")
	End If 
End If
If Request.Form("Preview") <> "" and Request.QueryString("CardNo") = "" Then 
strTo = Request.Form("C_To")
strToEmail = Request.Form("C_ToEmail")
strFrom = Request.Form("C_From")
strFromEmail = Request.Form("C_FromEmail")
strMessage = Request.Form("C_Message")
strUrl = Request.Form("C_Url")
If Session("MM_UserName") = "" Then 
 Response.Redirect("/ecards.asp?Msg=Please Login First")
End If 
End If 
%>
<!--#Include virtual="/assets/scripts/var.asp"-->
<html>
<head>
<title>Ramadaan.co.za - </title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript" type="text/JavaScript">
<!--

function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
//-->
</script>
<link href="/assets/scripts/ramadaan.css" rel="stylesheet" type="text/css">
<link href="/assets/scripts/menu.css" rel="stylesheet" type="text/css">
<meta content="For muslim users about during the month of ramadaan">
<meta description="For muslim users about during the month of ramadaan">
<meta keywords="Ramadaan, Ramadhaan, Ramadan, fasting, muslim, islamic, adhaan, software, eid, fitr, namaaz, prayer">
</head>
<body bgcolor="#D5DBBF" text="#333333" link="#003300" vlink="#999999" alink="#FF0000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" <% If Request.QueryString("Msg") <> "" And Request.QueryString("Deny") = "" Then response.Write("onLoad=alertmsg();")%>>
<div id="Layer3" style="position:absolute; left:0px; top:95px; width:779; height:100px; z-index:9; background-color: #FFFFFF; layer-background-color: #FFFFFF;"> 
  <table width="100%" border="0" class="text">
    <tr> 
      <td width="80%"><div align="right"><font size="3">To :</font></div></td>
      <td width="20%">
	  <%=strTo %>
        </td>
    </tr>
    <tr> 
      <td>
	  <img src="<%=strURL%>">
	  </td>
      <td>
	  <%=StrMessage%>
	  </td>
    </tr>
    <tr> 
      <td><div align="right">
          <font size="3">From :</font></div></td>
      <td>
	  <%=StrFrom%>
	  </td>
    </tr>
  </table>
  	  	<% if Request.Form("Preview") <> "" Then %>
		<div align="right">
          <form name="form1" method="POST" action="assets/scripts/sendcard.asp">
            <input name="C_To" type="hidden" id="C_To" value="<%=Request.Form("C_To") %>">
            <input name="C_ToEmail" type="hidden" id="C_ToEmail" value="<%=Request.Form("C_ToEmail") %>">
            <input name="C_From" type="hidden" id="C_From" value="<%=Request.Form("C_From") %>">
            <input name="C_FromEmail" type="hidden" id="C_FromEmail" value="<%=Request.Form("C_FromEmail") %>">
            <input name="C_URL" type="hidden" id="C_URL" value="<%=Request.Form("C_URL") %>">
            <input name="C_Message" type="hidden" id="C_Message" value="<%=Request.Form("C_Message") %>">
            <input name="Send" type="submit" class="fields" id="Send" value="Send">
            <input type="hidden" name="MM_insert" value="form1">
          </form></div>
		  <% End If %>
</div>
<table border="0" cellpadding="0" cellspacing="0" width="779">
  <!-- fwtable fwsrc="main.png" fwbase="main.gif" fwstyle="Dreamweaver" fwdocid = "742308039" fwnested="0" -->
  <tr> 
    <td><img src="/assets/images/spacer.gif" width="78" height="1" border="0" alt=""></td>
    <td><img src="/assets/images/spacer.gif" width="227" height="1" border="0" alt=""></td>
    <td><img src="/assets/images/spacer.gif" width="468" height="1" border="0" alt=""></td>
    <td><img src="/assets/images/spacer.gif" width="6" height="1" border="0" alt=""></td>
    <td><img src="/assets/images/spacer.gif" width="1" height="1" border="0" alt=""></td>
  </tr>
  <tr> 
    <td colspan="4"><img name="main_r1_c1" src="/assets/images/main_r1_c1.gif" width="779" height="5" border="0" alt=""></td>
    <td><img src="/assets/images/spacer.gif" width="1" height="5" border="0" alt=""></td>
  </tr>
  <tr> 
    <td rowspan="2" colspan="2"><a href="http://www.ramadaan.co.za"><img name="main_r2_c1" src="/assets/images/main_r2_c1.gif" width="305" height="65" border="0" alt=""></a></td>
    <td><!--#include virtual="/assets/scripts/banners.asp"--></td>
    <td rowspan="2"><img name="main_r2_c4" src="/assets/images/main_r2_c4.gif" width="6" height="65" border="0" alt=""></td>
    <td><img src="/assets/images/spacer.gif" width="1" height="60" border="0" alt=""></td>
  </tr>
  <tr> 
    <td><img name="main_r3_c3" src="/assets/images/main_r3_c3.gif" width="468" height="5" border="0" alt=""></td>
    <td><img src="/assets/images/spacer.gif" width="1" height="5" border="0" alt=""></td>
  </tr>
  <tr> 
    <td><img name="main_r4_c1" src="/assets/images/main_r4_c1.gif" width="78" height="22" border="0" alt=""></td>
    <td colspan="3" valign="top" bgcolor="#ffffff"><table border="0" cellpadding="0" cellspacing="0" width="700">
        <!-- fwtable fwsrc="timesbann.png" fwbase="timesban.gif" fwstyle="Dreamweaver" fwdocid = "742308039" fwnested="0" -->
        <tr> 
          <td><img src="/assets/images/spacer.gif" alt="" name="undefined_2" width="25" height="1" border="0"></td>
          <td><img src="/assets/images/spacer.gif" alt="" name="undefined_2" width="158" height="1" border="0"></td>
          <td><img src="/assets/images/spacer.gif" alt="" name="undefined_2" width="35" height="1" border="0"></td>
          <td><img src="/assets/images/spacer.gif" alt="" name="undefined_2" width="40" height="1" border="0"></td>
          <td><img src="/assets/images/spacer.gif" alt="" name="undefined_2" width="35" height="1" border="0"></td>
          <td><img src="/assets/images/spacer.gif" alt="" name="undefined_2" width="93" height="1" border="0"></td>
          <td><img src="/assets/images/spacer.gif" alt="" name="undefined_2" width="35" height="1" border="0"></td>
          <td><img src="/assets/images/spacer.gif" alt="" name="undefined_2" width="41" height="1" border="0"></td>
          <td><img src="/assets/images/spacer.gif" alt="" name="undefined_2" width="35" height="1" border="0"></td>
          <td><img src="/assets/images/spacer.gif" alt="" name="undefined_2" width="88" height="1" border="0"></td>
          <td><img src="/assets/images/spacer.gif" alt="" name="undefined_2" width="35" height="1" border="0"></td>
          <td><img src="/assets/images/spacer.gif" alt="" name="undefined_2" width="41" height="1" border="0"></td>
          <td><img src="/assets/images/spacer.gif" alt="" name="undefined_2" width="35" height="1" border="0"></td>
          <td><img src="/assets/images/spacer.gif" alt="" name="undefined_2" width="4" height="1" border="0"></td>
          <td><img src="/assets/images/spacer.gif" alt="" name="undefined_2" width="1" height="1" border="0"></td>
        </tr>
        <tr> 
          <td><img name="timesbann_RDate" src="/assets/images/times/timesbann_RDate<%=strRamDate%>.gif" width="25" height="20" border="0" alt=""></td>
          <td><img name="timesban_r1_c2" src="/assets/images/timesban_r1_c2.gif" width="158" height="20" border="0" alt=""></td>
          <td><img name="timesbann_jhb_s" src="/assets/images/times/timesbann_jhb_s<%=strRamDate%>.gif" width="35" height="20" border="0" alt=""></td>
          <td><img name="timesban_r1_c4" src="/assets/images/timesban_r1_c4.gif" width="40" height="20" border="0" alt=""></td>
          <td><img name="timesbann_jhb_i" src="/assets/images/times/timesbann_jhb_i<%=strRamDate%>.gif" width="35" height="20" border="0" alt=""></td>
          <td><img name="timesban_r1_c6" src="/assets/images/timesban_r1_c6.gif" width="93" height="20" border="0" alt=""></td>
          <td><img name="timesbann_dbn_s" src="/assets/images/times/timesbann_dbn_s<%=strRamDate%>.gif" width="35" height="20" border="0" alt=""></td>
          <td><img name="timesban_r1_c8" src="/assets/images/timesban_r1_c8.gif" width="41" height="20" border="0" alt=""></td>
          <td><img name="timesbann_dbn_i" src="/assets/images/times/timesbann_dbn_i<%=strRamDate%>.gif" width="35" height="20" border="0" alt=""></td>
          <td><img name="timesban_r1_c10" src="/assets/images/timesban_r1_c10.gif" width="88" height="20" border="0" alt=""></td>
          <td><img name="timesbann_cpt_s" src="/assets/images/times/timesbann_cpt_s<%=strRamDate%>.gif" width="35" height="20" border="0" alt=""></td>
          <td><img name="timesban_r1_c12" src="/assets/images/timesban_r1_c12.gif" width="41" height="20" border="0" alt=""></td>
          <td><img name="timesbann_cpt_i" src="/assets/images/times/timesbann_cpt_i<%=strRamDate%>.gif" width="35" height="20" border="0" alt=""></td>
          <td><img name="timesban_r1_c14" src="/assets/images/timesban_r1_c14.gif" width="4" height="20" border="0" alt=""></td>
          <td><img src="/assets/images/spacer.gif" alt="" name="undefined_2" width="1" height="20" border="0"></td>
        </tr>
      </table></td>
    <td><img src="/assets/images/spacer.gif" width="1" height="22" border="0" alt=""></td>
  </tr>
</table>
</body>
</html>
<%
If Request.QueryString("CardNo") <> "" Then 
rsCard.Close()
Set rsCard = Nothing
End If 
%>
