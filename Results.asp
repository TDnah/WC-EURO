<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="Stylesheet" href="Results.css.asp" />
<title>Thong tin chon doi</title>

<style type="text/css">
<!--
.TieuDe {color: #FFFF00;
	font-weight: bold;
	font-size: 30px;
}
-->
#Filter
{
    text-transform:none;
    color: #FFFF00;
}
	
#tblResults td
{
	 text-align:center;
	 border-right: solid thin #333333;
	 border-left: solid thin #333333;
}

#tblResults th
{
	 text-align:center;
	 border: solid thin #333333;
}
</style>
<script type="text/javascript" language="javascript">
    function doFilter(objSelect)
    {
        document.location = "Results.asp?MatchType=" + objSelect.value;
    }
</script>
</head>

<body>

<%@ Language=VBScript %>
<%Session.CodePage=65001%>
<!--#INCLUDE FILE="Gameheader.asp"-->
<%Dim MatchType, strWhere
MatchType = Request("MatchType") & ""
if MatchType="" then
    MatchType="2"
end if
strWhere = ""
if MatchType="2" then strWhere = " WHERE Matchs.Description Is null"
if MatchType="3" then strWhere = " WHERE Matchs.Description Is not null"
%>
<div class="Title"><span class="TieuDe">Bảng theo dõi tình hình chọn dội</span>
<div id="Filter">Hiển thị
<select name="MatchType" id="MatchTypeFilter" onchange="doFilter(this);">
    <option value="1" <%if MatchType="1" then Response.Write "Selected" %>>Tất cả các trận</option>
    <option value="2" <%if MatchType="2" then Response.Write "Selected"%>>Các trận chưa thi đấu</option>
    <option value="3" <%if MatchType="3" then Response.Write "Selected" %>>Các trận đã thi đấu</option>
</select>
</div>
</div>

<table background="Image/Bg.jpg" class="TableList" cellspacing="0" id="tblResults">
<%
  Dim Conn, rsGame, rsUser1, rsUser2,rsUser3, SQL, SQL1,SQL3
  dim check
  SQL = "SELECT * FROM qGamesToday " & strWhere & " Order by MatchID "
  Set Conn = Server.CreateObject("ADODB.Connection")
  Conn.Open strConnString
  Set rsUser1 = Server.CreateObject("ADODB.Recordset")
  Set rsUser2 = Server.CreateObject("ADODB.Recordset")
  Set rsUser3 = Server.CreateObject("ADODB.Recordset")
  Set RS = Server.CreateObject("ADODB.Recordset")
  RS.Open SQL, Conn
%>
	<%i=1%>
<%Do while not RS.EOF%>
	<tr>
		<th colspan="4"><font size="+1"><%=RS.Fields("Team1Name")%></font>
		 - Tỷ số chấp: <%=RS.Fields("Matchs.Ratio")%> 
		 | Thắng: <%=RS.Fields("Win")%>
		 | Hòa: <%=RS.Fields("Draw")%>
		 | Thua: <%=RS.Fields("Lose")%>
		 - <font size="+1"><%=RS.Fields("Team2Name")%></font></th>
	</tr>
	<tr>
        <th align="center" width="20%">Ngày giờ chọn</th>
        <th align="center" width="30%">Họ Tên</th>
        <th align="center" width="30%">Họ Tên</th>
        <th align="center" width="20%">Ngày giờ chọn</th>
    </tr>
<%
	SQL1="SELECT UserName,Users.UserID,UpdatedDate   FROM Users INNER JOIN GAMES on Users.UserID=Games.UserID WHERE MatchID=" &  RS.Fields("MatchID") & " AND TeamID='"  & RS.Fields("Matchs.Team1") & "' ORDER BY UpdatedDate"
	rsUser1.Open SQL1, Conn
	SQL2="SELECT UserName,Users.UserID, UpdatedDate FROM Users INNER JOIN GAMES on Users.UserID=Games.UserID WHERE MatchID=" & RS.Fields("MatchID") & " AND TeamID='"  & RS.Fields("Matchs.Team2") & "' ORDER BY UpdatedDate"
	rsUser2.Open SQL2, Conn
	SQL3="SELECT UserName ,Users.UserID FROM Users INNER JOIN GAMES on Users.UserID=Games.UserID WHERE MatchID=" & RS.Fields("MatchID") & " AND TeamID='Chua' ORDER BY UserName"
	rsUser3.Open SQL3, Conn
	%>	                         			
	<%Do while (not rsUser1.EOF) or (not rsUser2.EOF)%>
		  <tr>
			<%if not rsUser1.EOF then%>
				<td><font size="-1"><%=rsUser1.Fields("UpdatedDate")%>&nbsp;</font></td>
				<td class="<%=rsUser1.Fields("UserID")%>"><span class="<%=replace(replace(replace(rsUser1.Fields("UserID")," ",""),"@",""),".","")%>_thinking">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span><%=rsUser1.Fields("UserName")%>&nbsp;<span class="<%=replace(replace(replace(rsUser1.Fields("UserID")," ",""),"@",""),".","")%>">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
				<%rsUser1.movenext%>
			<%else%>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
			<%end if%>

			<%if not rsUser2.EOF then%>
				<td class="<%=rsUser2.Fields("UserID")%>"><span class="<%=replace(replace(replace(rsUser2.Fields("UserID")," ",""),"@",""),".","")%>_thinking">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span><%=rsUser2.Fields("UserName")%>&nbsp;<span class="<%=replace(replace(replace(rsUser2.Fields("UserID")," ",""),"@",""),".","")%>">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
				<td><font size="-1"><%=rsUser2.Fields("UpdatedDate")%>&nbsp;</font></td>
				<%rsUser2.movenext%>
			<%else%>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
			<%end if%>
		</tr>
	<%Loop%>	

    <%if (not rsUser3.EOF) then %>
	<tr>
		<th colspan="4" align="center"><b>Danh sách những người chắc chắn thua:</b></th>
	</tr>
	<%End if %>
	<%Do while (not rsUser3.EOF)%>
		<tr>
			<td colspan="4" class="<%=replace(rsUser3.Fields("UserID")," ","")%>"><span class="<%=replace(replace(replace(rsUser3.Fields("UserID")," ",""),"@",""),".","")%>_thinking">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span><%=rsUser3.Fields("UserName")%>&nbsp;<span class="<%=replace(replace(replace(rsUser3.Fields("UserID")," ",""),"@",""),".","")%>">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
		</tr>
	<%rsUser3.movenext%>
	<%loop%>
	<%rsUser1.Close%>
	<%rsUser2.Close%>
	<%rsUser3.Close%>
	<%i=i+1%>
	<%RS.MoveNext() %>
<%Loop%>
</table>

</body>
</html>