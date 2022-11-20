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
	  Dim Conn, sSQL, MatchID, UserID, TeamID, pScore1, pScore2, pRatio, pWin, pDraw, pLose, pDesc, RecAffected, i
    Set Conn = Server.CreateObject("ADODB.Connection")
	  Conn.Open strConnString
	  UserID=session("UserID")
    for i=1 to 64
      MatchID=Request.Form("MatchID" & i)
      pRatio=Request.Form("InputRatio" & i)
      pScore1=Request.Form("InputScore1" & i)
      pScore2=Request.Form("InputScore2" & i)
      pDesc=Request.Form("InputDesc" & i)
      pWin=Request.Form("InputWin" & i)
			pDraw=Request.Form("InputDraw" & i)
			pLose=Request.Form("InputLose" & i)
      if pDesc<>"" then
        sSQL="Update Matchs SET Ratio=" & pRatio & ", Score1=" & pScore1 & ", Score2=" & pScore2 & ", Win=" & pWin & ", Draw=" & pDraw & ", Lose=" & pLose & ", Matchs.Description = 'Pass' WHERE MatchID=" & MatchID & ""
      else
        sSQL="Update Matchs SET Ratio=" & pRatio & ", Score1=" & pScore1 & ", Score2=" & pScore2 & ", Win=" & pWin & ", Draw=" & pDraw & ", Lose=" & pLose & ", Matchs.Description = NULL WHERE MatchID=" & MatchID & ""
      end if
      conn.execute sSQL,RecAffected
			if RecAffected > 0 then
				Response.Write("<H3>Game " & sSQL & ": Updated successfully</H3>")
			end if
    next
%>

</body>

</html>