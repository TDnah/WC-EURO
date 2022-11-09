<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Lich thi dau</title>
<style type="text/css">
<!--
.TieuDe {	color: #FFFF00;
	font-weight: bold;
	font-size: 30px;
}
-->
</style>
</head>

<body>
<%@ Language=VBScript %>
<%Session.CodePage=65001%>
<!--#INCLUDE FILE="Gameheader.asp"-->

<div class="Title"><span class="TieuDe">Lịch Thi Đấu </span></div>
<table background="Image/Bg.jpg" class="TableList">
  <tr>
	<th>Trận</th>
	<th>Ngày giờ</th>
	<th>Đội 1</th>
	<th>Đội 2</th>
	<th>Tỷ lệ chấp</th>
	<th>Thắng</th>
	<th>Hòa</th>
	<th>Thua</th>
  </tr>
  <%
  Dim Conn, RS, SQL
  dim check
  SQL = "SELECT * FROM qFixtures Order by [Date]"
  Set Conn = Server.CreateObject("ADODB.Connection")
  'ms access dsnless connection
  Conn.Open strConnString
  Set RS = Server.CreateObject("ADODB.Recordset")
  RS.Open SQL, Conn
  %>
  <form METHOD="post" ACTION="UpdateGames.asp">
	<%i=1%>
	<%Do while not RS.EOF%>
	<tr>
	  <td align="center"><%=i%> </td>
	  <td><%=day(RS.Fields("Date")) & "/" & month(RS.Fields("Date")) & "/" & year(RS.Fields("Date")) & " " & FormatDateTime(RS.Fields("Date"),3)%></td>
	  <td><%=RS.Fields("Team1Name")%></td>
	  <td><%=RS.Fields("Team2Name")%></td>
	  <td align="center"><%=RS.Fields("Ratio")%></td>
	  <td align="right"><%=RS.Fields("Win")%></td>
	  <td align="right"><%=RS.Fields("Draw")%></td>
	  <td align="right"><%=RS.Fields("Lose")%></td>
	</tr>
	<tr>
	  <%i=i+1%>
	  <%RS.MoveNext() %>
	  <%Loop%>
	  </form>
</table>

</body>
</html>