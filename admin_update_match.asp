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
	Dim Conn, sSQL, MatchID, UserID, TeamID, pTeam1, pTeam2, RecAffected, i
    Set Conn = Server.CreateObject("ADODB.Connection")
    Conn.Open strConnString
    UserID=session("UserID")
    for i=1 to 51
        MatchID=Request.Form("MatchID" & i)
        pTeam1=Request.Form("InputTeam1" & i)
        pTeam2=Request.Form("InputTeam2" & i)
        sSQL="Update Matchs SET Team1='" & pTeam1 & "', Team2='" & pTeam2 & "' WHERE MatchID=" & MatchID & ""
        Response.Write("<H3>Game " & sSQL & ": Updated successfully</H3>")
        conn.execute sSQL,RecAffected
        if RecAffected > 0 then
            Response.Write("<H3>Game " & sSQL & ": Updated successfully</H3>")
        end if
    next
%>

</body>

</html>