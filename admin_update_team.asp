<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Update data</title>
</head>

<body>

<%@ Language=VBScript %>
<%Session.CodePage=65001%>

<!--#INCLUDE FILE="Gameheader.asp"-->

<%
	Dim Conn, sSQL, pTeamId, pTeamName, RecAffected, i
    Set Conn = Server.CreateObject("ADODB.Connection")
    Conn.Open strConnString
    UserID=session("UserID")
    for i=1 to 51
        pTeamId=Request.Form("iTeamID" & i)
        pTeamName=Request.Form("iTeamName" & i)
        sSQL="Update Teams SET TeamName='" & pTeamName & "' WHERE TeamID=" & pTeamId & ""
        Response.Write("<H3>Team " & sSQL & ": Updated successfully</H3>")       
    next
%>

</body>

</html>