<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta http-equiv="Content-Language" content="en-us">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Cap nhat doi bong</title>
<style type="text/css">
<!--
.TieuDe {color: #FFFF00;
	font-weight: bold;
	font-size: 30px;
}
.style1 {
	color: #993300;
	font-weight: bold;
}
.style2 {
	color: #996600;
	font-weight: bold;
}
-->
</style>
</head>

<body>

<%@ Language=VBScript%>
<%Session.CodePage=65001%>
<!--#INCLUDE FILE="Gameheader.asp"-->

  <div> Server time: <%=FormatDateTime(now())%> </div> 
  <div class="Title"><span class="TieuDe">Các trận đấu diễn ra trong các ngày sắp tới</span></div>

<table background="Image/Bg.jpg" class="TableList">
<br/>
 <%
  Dim Conn, RS, SQL
  SQL = "SELECT * FROM qMatchsInfo ORDER BY MatchID"
  Set Conn = Server.CreateObject("ADODB.Connection")
  Conn.Open strConnString
  Set RS = Server.CreateObject("ADODB.Recordset")
  RS.Open SQL, Conn
%>


<form METHOD="post" ACTION="admin_update.asp">
	<%i=1%>
	<%Do while not RS.EOF%>
		<tr>
			<td align="center"><%=RS.Fields("MatchID")%></td>
			<td><%=day(RS.Fields("Matchs.Date")) & "/" & month(RS.Fields("Matchs.Date")) & "/" & year(RS.Fields("Matchs.Date")) & " " & FormatDateTime(RS.Fields("Matchs.Date"),3)%></td>			
            <td align="center"><%=RS.Fields("Team1")%></td>
			<td><%=RS.Fields("Team1Name")%></td>		
            <td align="center"><%=RS.Fields("Team2")%></td>
			<td><%=RS.Fields("Team2Name")%></td>
            <td>
                <select>
                    <%
                        Dim Conn1, RS1, SQL1
                        dim check
                        SQL1 = "SELECT * FROM qTeams ORDER BY MatchID"
                        Set Conn1 = Server.CreateObject("ADODB.Connection")
                        Conn1.Open strConnString
                        Set RS1 = Server.CreateObject("ADODB.Recordset")
                        RS1.Open SQL1, Conn1
                    %>
                    <option value=""></option> 
                    <%Do while not RS1.EOF%>
                        <option value="<%=RS1.Fields("Teams.TeamID")%>"><%=RS1.Fields("Teams.TeamName")%></option>
                    <%RS1.MoveNext() %>
                    <%Loop%>
                </select>
            </td>
            <td>-</td>
		</tr>
		<input type="hidden" id="text1" name="MatchID<%=i%>" Value="<%=RS.Fields("MatchID")%>">
		<tr>
    <%i=i+1%>
    <%RS.MoveNext() %>
    <%Loop%>
</table>
<br/>
<table width="100%">
	<tr>
		<td align="center">
		<input type="submit" value="  Submit  " id="submit1" name="submit1"></td>
		</td>
	</tr>
</table>
</form>

</body>
</html>