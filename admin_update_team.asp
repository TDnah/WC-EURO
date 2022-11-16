<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta http-equiv="Content-Language" content="en-us">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Cap nhat doi</title>
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
.sidenav {
  height: 100%;
  width: 250px;
  position: fixed;
  z-index: 1;
  top:0;
  left: 0;
  background-color: #111;
  overflow-x: hidden;
  padding-top: 20px;
}

.main {
  margin-left: 250px; /* Same as the width of the sidenav */
  font-size: 28px; /* Increased text to enable scrolling */
  padding: 0px 10px;
}
-->
</style>
</head>

<body>

<%@ Language=VBScript%>
<%Session.CodePage=65001%>

<div class="container" width="100%">
    <div class="main">
        <!--#INCLUDE FILE="Gameheader.asp"-->
    </div>
    <div class="sidenav">
        <table>
            <tr>
                <th colSpan="2"><p align="center"><font color="white">Mã đội</th>
                <th colSpan="2"><p align="center"><font color="white">Tên đội</th>
                <th colSpan="2"><p align="center"><font color="white">Mô tả</th>
            </tr>
            <%
            Dim Conn, RS, SQL
            SQL = "SELECT * FROM qTeams WHERE Description Is not null ORDER BY Description"
            Set Conn = Server.CreateObject("ADODB.Connection")
            Conn.Open strConnString
            Set RS = Server.CreateObject("ADODB.Recordset")
            RS.Open SQL, Conn
            %>
            <% DO while not RS.EOF %>
            <tr>
                <td colSpan="2">
                    <p align="center" style="color:White;"><%=RS.Fields("TeamID")%></p>
                </td>
                <td colSpan="2">
                    <p align="center" style="color:White;"><%=RS.Fields("TeamName")%></p>
                </td>
                <td colSpan="2">
                    <p align="center" style="color:White;"><%=RS.Fields("Description")%></p>
                </td>
            </tr>
            <% RS.Movenext() %>
            <% Loop %>
            <%
                RS.Close
                Conn.Close
            %>
        </table>
    </div>
    <br/>
    <div class="main">
        <table>
            <tr>
                <th>Trận</th>
                <th>Ngày giờ </th>
                <th colSpan="2"><p align="center"><font color="white">Đội 1</th>
                <th colSpan="2"><p align="center"><font color="white">Đội 2</th>
                <th colSpan="2"><p align="center"><font color="white">Cập nhật đội 1</th>
                <th colSpan="2"><p align="center"><font color="white">Cập nhật đội 2</th>
            </tr>
            <%
                SQL = "SELECT * FROM qMatchsInfo WHERE MatchID>=49 ORDER BY MatchID"
                Set Conn = Server.CreateObject("ADODB.Connection")
                Conn.Open strConnString
                Set RS = Server.CreateObject("ADODB.Recordset")
                RS.Open SQL, Conn
            %>

        <form METHOD="post" ACTION="admin_update_match.asp">
            <%i=49%>
            <%Do while not RS.EOF%>
                <tr>
                    <td align="center"><%=RS.Fields("MatchID")%></td>
                    <td><%=day(RS.Fields("Matchs.Date")) & "/" & month(RS.Fields("Matchs.Date")) & "/" & year(RS.Fields("Matchs.Date")) & " " & FormatDateTime(RS.Fields("Matchs.Date"),3)%></td>		
                    <td colSpan="2">
                        <p align="center"><%=RS.Fields("Team1")%></p>
                        <p align="center"><%=RS.Fields("Team1Name")%></p>
                    </td>
                    <td colSpan="2">
                        <p align="center"><%=RS.Fields("Team2")%></p>
                        <p align="center"><%=RS.Fields("Team2Name")%></p>
                    </td>
                    <td colSpan="2" align="center">
                        <input id="team1" type="text" style="text-transform: uppercase;" name="InputTeam1<%=i%>" value="<%=RS.Fields("Team1")%>">
                    </td>
                    <td colSpan="2" align="center">
                        <input id="team2" type="text" style="text-transform: uppercase;" name="InputTeam2<%=i%>" value="<%=RS.Fields("Team2")%>">
                    </td>
                </tr>
                <input type="hidden" id="text1" name="MatchID<%=i%>" Value="<%=RS.Fields("MatchID")%>">
                <tr>
            <%i=i+1%>
            <%RS.MoveNext() %>
            <%Loop%>
        </table>
        <br/>
        <input type="submit" value="  Submit  " id="submit1" name="submit1">
        </form>
    </div>
</div>

</body>
</html>