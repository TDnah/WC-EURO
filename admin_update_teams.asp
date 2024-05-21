<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Danh Sách Các Đội Bóng</title>
<style type="text/css">
<!--
.style1 {
	font-size: 30px;
	color: #FFFF00;
	font-weight: bold;
}
.TieuDe {color: #FFFF00;
	font-weight: bold;
	font-size: 30px;
}
body {
	background-image: url();
}
-->
</style>
</head>

<body >
<%@ Language=VBScript %>
<%Session.CodePage=65001%>

<!--#INCLUDE FILE="Gameheader.asp"-->

  <div class="Title"><span class="TieuDe">Danh Sách Các Đội Bóng</span></div>
    <%   Dim Conn, RS, SQL
    SQL = "SELECT TeamID,TeamName,Description FROM Teams where Description<>Null ORDER BY Index,TeamID"
    Set Conn = Server.CreateObject("ADODB.Connection")
    Conn.Open strConnString
    Set RS = Server.CreateObject("ADODB.Recordset")
    RS.Open SQL, Conn
    %>
    <table background="Image/Bg.jpg"  class="TableList">
        <tr bgcolor="seagreen" color="white" >
            <th>Mã số</th>
            <th>Tên đội</th>
            <th>Ghi chú</th>
        </tr>	
        <form METHOD="post" ACTION="admin_update_team.asp"></form>
            <%i=1%>
            <%Do while not RS.EOF%>
            <tr>
                    <td>
                        <p align="center" ><%=RS.Fields("TeamID")%></p>
                    </td>
                    <td>
                        <input id="teamName" type="text" style="text-transform: uppercase; width: 100%;" name="iTeamName<%=i%>" value="<%=RS.Fields("TeamName")%>">               
                    </td>
                    <td>
                        <p align="center"><%=RS.Fields("Description")%></p>
                    </td>
                    <input type="hidden" id="text1" name="iTeamID<%=i%>" Value="<%=RS.Fields("TeamID")%>">
            </tr>
            <%i=i+1%>
            <%RS.MoveNext() %>
            <%Loop%>
    </table>
    <br/>
    <input type="submit" value="  Submit  " id="submit1" name="submit1">
    </form>
</body>
</html>