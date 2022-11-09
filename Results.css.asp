<%@ Language=VBScript %>
<%Session.CodePage=65001%>
<!--#INCLUDE FILE="config.asp"-->
<%
Dim Conn, rs, SQL
SQL = "SELECT * FROM qResultsAll_Crosstab1 WHERE TotalMoney = (SELECT Max(TotalMoney) FROM qResultsAll_Crosstab1)"
Set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open strConnString

Set rs = Server.CreateObject("ADODB.Recordset")
RS.Open SQL, Conn

Do while not RS.EOF
    Response.write "span." & replace(replace(replace(RS("UserID")," ",""),"@",""),".","") & " {" & vbNewLine
    Response.write "background: url(flag.gif) no-repeat center;" & vbNewLine
    Response.write "}" & vbNewLine
    RS.MoveNext()
Loop

Response.write "td." & replace(replace(replace(Session("UserID")," ",""),"@",""),".","") & " {" & vbNewLine
Response.write "color: red;" & vbNewLine
Response.write "font-weight: bold;" & vbNewLine
Response.write "}" & vbNewLine

Response.write "span." & replace(replace(replace(Session("UserID")," ",""),"@",""),".","") & "_thinking {" & vbNewLine
Response.write "background: url(thinking.gif) no-repeat center;" & vbNewLine
Response.write "}" & vbNewLine

%>