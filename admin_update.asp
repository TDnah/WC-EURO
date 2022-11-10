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
        pScore1=Request.Form("InputScore1" & i)
        pScore2=Request.Form("InputScore2" & i)
        pRatio=Request.Form("InputRatio" & i)
        if User_ID = "admin"
            if TeamID<>"" then
                MatchID=Request.Form("MatchID" & i)
                sCriteria="TeamID<>'" & TeamID & "' And MatchID=" & MatchID & " AND UserID='" & UserID & "'"
                sSQL="Update Games SET Score1='" & pScore1 & "', Score2='" & pScore2 &"', Ratio='" & pRatio &"', UpdatedDate='" & convDate & "' WHERE " & sCriteria  
                conn.execute sSQL,RecAffected
                if RecAffected > 0 then
                    Response.Write("<H3>Game " & i & ": Updated successfully</H3>")
                end if
            end if
                next
        end   
	   Response.Redirect("Results.asp")
%>

</body>

</html>