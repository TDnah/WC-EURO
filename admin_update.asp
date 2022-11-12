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
      
	  Dim Conn, sSQL, MatchID, UserID, TeamID, pScore1, pScore2, pRatio, sCriteria, RecAffected
	  Dim convDate
      convDate = now()
      convDate = DateAdd("h", 14, convDate)
      Set Conn = Server.CreateObject("ADODB.Connection")
	  Conn.Open strConnString
	  UserID=session("UserID")
	  for i=1 to Request.Form.count \ 2
        pScore1=Request.Form("InputScore1" & pS1)
        pScore2=Request.Form("InputScore2" & pS2)
        pRatio=Request.Form("InputRatio" & pR)
        MatchID=Request.Form("MatchID" & i)
        sCriteria="MatchID='" & MatchID & "'"
        sSQL="Update Matchs SET Score1='" & pScore1 & "', Score2='" & pScore2 &"', Ratio='" & pRatio &"' WHERE " & sCriteria  
        conn.execute sSQL,RecAffected
        if RecAffected > 0 then
            Response.Write("<H3>Game " & i & ": Updated successfully</H3>")
        end if
        next  
	   Response.Redirect("Results.asp")
%>

</body>

</html>