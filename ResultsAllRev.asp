<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Bang thanh tich ca nhan</title>
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

.style2 {
	color: #FFFFFF;
	font-weight: bold;
	font-size: 20px;
}
-->
</style>
</head>
<body>

<%@ Language=VBScript %>
<%Session.CodePage=65001%>
<!--#INCLUDE FILE="Gameheader.asp"-->
<%   Dim Conn, RS, SQL
  SQL = "SELECT UserName, TotalMoney, TotalWinMatch, TotalDrawMatch, TotalLoseMatch FROM qResultsAll_Crosstab1 INNER JOIN Users ON qResultsAll_Crosstab1.UserID=Users.UserID ORDER BY TotalMoney DESC"
  SQL1 = "SELECT sum(TongST) FROM qResultsAll_Crosstab"
  Set Conn = Server.CreateObject("ADODB.Connection")
  Conn.Open strConnString
  Set RS = Server.CreateObject("ADODB.Recordset")
  Set RS1 = Server.CreateObject("ADODB.Recordset")
  RS.Open SQL, Conn
  RS1.Open SQL1, Conn
%>
  
  <div class="Title"><span class="TieuDe">Bảng thành tích cá nhân <br/>

    <%if not RS.EOF then %>
        <span class="style2">Tổng số tiền thu đươc: <%=FormatNumber(RS1.Fields(0),0)%></span>
    <%End if %>
  </div>

<table background="Image/Bg.jpg" class="TableList">
<tr>
		<th>Hạng</th>
		<th>Người tham gia</th>
		<th>Tổng sồ tiền đóng góp</th>
		<th>Số trận thắng</th>
		<th>Số trận hòa</th>
		<th>Số trận thua</th>
	</tr>	
    <%stt=1%>
	<%Do while not RS.EOF%>
	<tr>
	<td align="center"><%=stt%>&nbsp;</td>
    <%For i=0 to 4%>
        <%if i>0 then %>
		    <td align="center"><%=FormatNumber(RS.Fields(i),0)%>&nbsp;</td>
		<%Else %>
		    <td align="center"><%=RS.Fields(i)%>&nbsp;</td>
		<%End if %>
	<%Next%> 
	</tr>	
	<%stt=stt+1%>
	<%RS.MoveNext() %>
<%Loop%>
</table>
</body>
</html>