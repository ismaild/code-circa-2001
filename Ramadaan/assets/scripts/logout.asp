<%
Session.Contents.Remove("MM_Username")
Session.Contents.Remove("MM_UserAuthorization")
Session.Contents.Remove("IM_FirstName")
Response.Redirect("/index.asp")
%>