﻿<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta http-equiv="Content-Language" content="en-us">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Chon doi</title>
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
	<th> Thắng </th>
	<th> Hòa </th>
	<th> Thua </th>
  </tr>
  <%
  Dim Conn, RS, SQL
  dim check
  SQL = "SELECT * FROM qMatchsToday WHERE UserID='" & Session("UserID") & "' ORDER BY Matchs.Date ASC, MatchID"
  Set Conn = Server.CreateObject("ADODB.Connection")
  Conn.Open strConnString
  Set RS = Server.CreateObject("ADODB.Recordset")
  RS.Open SQL, Conn
%>

<form METHOD="post" ACTION="Capnhatkq.asp">
<%i=1%>
<%Do while not RS.EOF%>
	<tr>
		<td align="center"><%=RS.Fields("MatchID")%></td>
		<td><%=day(RS.Fields("Matchs.Date")) & "/" & month(RS.Fields("Matchs.Date")) & "/" & year(RS.Fields("Matchs.Date")) & " " & FormatDateTime(RS.Fields("Matchs.Date"),3)%></td>
		<%if RS.Fields("TeamID")=RS.Fields("Team1") then %>
			<td align="center"><input id="radio" type="radio" name="SelectedTeam<%=i%>" value="<%=RS.Fields("Team1")%>" checked></td>
		<%else%>
			<td align="center"><input id="radio" type="radio" name="SelectedTeam<%=i%>" value="<%=RS.Fields("Team1")%>"></td>
		<%end if%>
		<td><%=RS.Fields("Team1Name")%></td>
		<%if RS.Fields("TeamID")=RS.Fields("Team2") then %>
			<td align="center"><input id="radio" type="radio" name="SelectedTeam<%=i%>" value="<%=RS.Fields("Team2")%>" checked></td>
		<%else%>
			<td align="center"><input id="radio" type="radio" name="SelectedTeam<%=i%>" value="<%=RS.Fields("Team2")%>"></td>
		<%end if%> 
		<td><%=RS.Fields("Team2Name")%></td>
		<td align="center"><%=RS.Fields("Ratio")%></td>
		<td align="right"><%=RS.Fields("Matchs.Win")%></td>
		<td align="right"><%=RS.Fields("Matchs.Draw")%></td>
		<td align="right"><%=RS.Fields("Matchs.Lose")%></td>
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
  <input type="submit" value="  Chọn  " id="submit1" name="submit1"></td>
  </td>
</tr>
</table>
</form>


</body>
</html>