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
  SQL = "SELECT TeamID,TeamName,Description FROM Teams where Description<>Null ORDER BY Description"
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
  
  	<%Do while not RS.EOF%>
    <tr align="left" >
<%For i=0 to 2%>
	  		<%if i=0 then%>
	      		<td nowrap><%=RS.Fields(i)%>&nbsp;</td>
	  		<%else%>
	      		<td nowrap ><%=RS.Fields(i)%>&nbsp;</td>
     		<%end if%>
	  <%Next%> 
	</tr>	
 	<%RS.MoveNext() %>
  	<%Loop%>
   <tr>
</table>

</body>
</html>