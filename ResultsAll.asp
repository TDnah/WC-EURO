<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Bang phong than</title>
<style type="text/css">
<!--
.TieuDe {
	color: #FFFF00;
	font-weight: bold;
	font-size: 30px;
}
.style2 {
	color: #FFFFFF;
	font-weight: bold;
	font-size: 20px;

}
#ResultTable td
    {
       border-right:thin solid Silver;
       border-bottom:thin dotted Silver;
       font-size:12px;
    }
    
    #ResultTable th
    {
       border-right:thin solid Silver;
       border-top:thin solid Silver;
       border-bottom:thin solid Silver;
       font-size:12px;
    }
    
    #ResultTable
    {
       border-left:thin solid Silver;
       border-bottom:thin double Silver;
    }
    
    #left
    {
    	float:left;
    	width:350px;
    }
    
    #right
    {
    	/*float:right;
    	width:655px;
    	left:355px;
    	position:absolute; 
	*/
    	overflow: scroll;
    	z-index:-1;
    }
    
-->
</style>

<script language="javascript"; type="text/javascript">
    function detemineClientWidth() {
        var x = 0;
        var y = 0;
        if (self.innerHeight) // all except Explorer
        {
            x = self.innerWidth;
            y = self.innerHeight;
        }
        else if (document.documentElement && document.documentElement.clientHeight)
        // Explorer 6 Strict Mode
        {
            x = document.documentElement.clientWidth;
            y = document.documentElement.clientHeight;
        }
        else if (document.body) // other Explorers
        {
            x = document.body.clientWidth;
            y = document.body.clientHeight;
        }
        return x;
    }
    
    function setContentWidth()
    {
        var x = detemineClientWidth();
        var content = document.getElementById("right");
        content.style.width = x - 380 + "px";
        
    }
</script>
</head>

<body>
<%@ Language=VBScript %>
<%Session.CodePage=65001%>
<!--#INCLUDE FILE="Gameheader.asp"-->
<%   Dim Conn, RS,RS1,SQL1, SQL
  SQL = "SELECT * FROM qResultsAll_Crosstab"
  SQL1 = "SELECT sum(TongST) FROM qResultsAll_Crosstab"
  Set Conn = Server.CreateObject("ADODB.Connection")
  Conn.Open strConnString
  Set RS = Server.CreateObject("ADODB.Recordset")
  Set RS1 = Server.CreateObject("ADODB.Recordset")
  RS.Open SQL, Conn
  RS1.Open SQL1, Conn
 %>
 
<div class="Title"><span class="TieuDe">Bảng phong thần </span><br/>

<%If not RS.EOF then%>
   <span class="style2">Tổng số tiền thu đươc: <%=FormatNumber(RS1.Fields(0),0)%></span></div>
    <div id="left">
    <table background="Image/Bg.jpg" class="TableList" id="ResultTable"  cellspacing="0" cellpadding="0">
        <tr>
		    <th>Trận</th>
		    <th NOWRAP>Đội - Tỷ lệ chấp - Kết quả</th>
		    <th NOWRAP>&nbsp;&nbsp;Tổng số tiền&nbsp;&nbsp;</th>
	    </tr>	

	    <%Do while not RS.EOF%>
		    <tr>
            <%For i=0 to 2%>
                <%If i=2 then%>
                    <td align="center" height="20" NOWRAP><b>
                    <% If RS.Fields(i) & "" = "" Then 
                            Response.write "&nbsp;"
                       else
                            Response.write FormatNumber(RS.Fields(i),0)
                    End If%> </b></td>
                <%Else %>
                    <td align="center" height="20" NOWRAP><b><%=RS.Fields(i)%></b></td>
                <%End if %>
		    <%Next%> 
		    </tr>	
		    <%RS.MoveNext() %>
	    <%Loop%>
	    <tr>
	        <td>&nbsp;</td>
	        <td align="right"><b>Tổng:&nbsp;&nbsp;</b></td>
	        <td align="center"><b><%=FormatNumber(RS1.Fields(0),0)%></b></td>
	    </tr>
    </table>
    </div>
    <%
    RS.close() 
    RS.Open SQL, Conn
    %>
    <div id="right">
    <table background="Image/Bg.jpg" class="TableList" id="ResultTable" cellspacing="0" cellpadding="0">
        <tr>
		    <%For i=3 to RS.Fields.Count -1%>
		        <th NOWRAP>&nbsp;&nbsp;<%=RS.Fields(i).Name%>&nbsp;&nbsp;</th>
		    <%Next%> 
	    </tr>	

	    <% Dim Tong()
	    ReDim Preserve Tong(RS.Fields.Count)
	    Do while not RS.EOF%>
		    <tr>
            <%For i=3 to RS.Fields.Count -1%>
                <td align="center" height="20" NOWRAP>
                    <% If RS.Fields(i) & "" = "" Then %>
                            &nbsp;
                       <%else%>
                            <%=FormatNumber(RS.Fields(i),0)%>
                            <%Tong(i) = Tong(i) + RS.Fields(i)%>            
                    <%End If%>
                </td>
                
		    <%Next%> 
		    </tr>	
		    <%RS.MoveNext() %>
	    <%Loop%>
	    
	    <tr>
		    <%For i=3 to RS.Fields.Count -1%>
		        <td align="center"><b><%=FormatNumber(Tong(i),0)%></b></td>
		    <%Next%> 
	    </tr>	
	    
    </table>
    </div>
<%End if %>
</body>
</html>