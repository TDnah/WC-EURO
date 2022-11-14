<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta http-equiv="Content-Language" content="en-us">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Cap nhat ty so</title>
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
  <tr>
	<th>Trận</th>
	<th>Ngày giờ </th>
	<th colSpan="2"><p align="center"><font color="white">Đội 1</th>
	<th colSpan="2"><p align="center"><font color="white">Đội 2</th>
	<th>Tỷ số chấp</th>
	<th>Tỷ số 1</th>
	<th>Tỷ số 2</th>
	<th>Thắng</th>
	<th>Hoà</th>
	<th>Thua</th>
	<th>Đã đá</th>
  </tr>
  <%
  Dim Conn, RS, SQL
  dim check
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
			<td align="center"><input id="ratio" type="number" step=any name="InputRatio<%=i%>" value="<%=RS.Fields("Matchs.Ratio")%>"></td>
			<td align="center"><input id="score1" type="number" name="InputScore1<%=i%>" value="<%=RS.Fields("Matchs.Score1")%>"></td>
			<td align="center"><input id="score2" type="number" name="InputScore2<%=i%>" value="<%=RS.Fields("Matchs.Score2")%>"></td>
			<td align="right"><input id="win" type="number" name="InputWin<%=i%>" value="<%=RS.Fields("Win")%>"></td>
			<td align="right"><input id="draw" type="number" name="InputDraw<%=i%>" value="<%=RS.Fields("Draw")%>"></td>
			<td align="right"><input id="lose" type="number" name="InputLose<%=i%>" value="<%=RS.Fields("Lose")%>"></td>
			<td align="right"><input id="desc" type="text" name="InputDesc<%=i%>" value="<%=RS.Fields("Description")%>"></td>
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