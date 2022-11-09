<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Update games</title>
</head>

<body>

<%@ Language=VBScript %>
<%Session.CodePage=65001%>

<!--#INCLUDE FILE="Gameheader.asp"-->

<%
      
	  Dim Conn, sSQL, MatchID, UserID, TeamID, sCriteria, RecAffected
	  Dim convDate
      convDate = now()
      convDate = DateAdd("h", 14, convDate)
      Set Conn = Server.CreateObject("ADODB.Connection")
	  Conn.Open strConnString
	  UserID=session("UserID") 
	  for i=1 to Request.Form.count \ 2
		TeamID=Request.Form("SelectedTeam" & i)
		if TeamID<>"" then
			MatchID=Request.Form("MatchID" & i)
			sCriteria="TeamID<>'" & TeamID & "' And MatchID=" & MatchID & " AND UserID='" & UserID & "' AND MatchID IN (SELECT MatchID FROM Matchs WHERE Date>Now()+#12/30/1899 14:0:0#)"
			sSQL="Update Games SET TeamID='" & TeamID & "', UpdatedDate='" & convDate & "' WHERE " & sCriteria 
			conn.execute sSQL,RecAffected
			if RecAffected > 0 then
				Response.Write("<H3>Game " & i & ": Updated successfully</H3>")
			end if
		end if
	   next
	   
	   Response.Redirect("Results.asp")
%>

</body>

</html>