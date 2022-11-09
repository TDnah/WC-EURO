<%
'#################################################################################
'## Snitz Forums 2000 v3.4.06
'#################################################################################
'## Copyright (C) 2000-06 Michael Anderson, Pierre Gorissen,
'##                       Huw Reddick and Richard Kinser
'##
'## This program is free software; you can redistribute it and/or
'## modify it under the terms of the GNU General Public License
'## as published by the Free Software Foundation; either version 2
'## of the License, or (at your option) any later version.
'##
'## All copyright notices regarding Snitz Forums 2000
'## must remain intact in the scripts and in the outputted HTML
'## The "powered by" text/logo with a link back to
'## http://forum.snitz.com in the footer of the pages MUST
'## remain visible when the pages are viewed on the internet or intranet.
'##
'## This program is distributed in the hope that it will be useful,
'## but WITHOUT ANY WARRANTY; without even the implied warranty of
'## MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'## GNU General Public License for more details.
'##
'## You should have received a copy of the GNU General Public License
'## along with this program; if not, write to the Free Software
'## Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.
'##
'## Support can be obtained from our support forums at:
'## http://forum.snitz.com
'##
'## Correspondence and Marketing Questions can be sent to:
'## manderson@snitz.com
'##
'#################################################################################
%>
<!--#INCLUDE FILE="config.asp" -->
<!--#INCLUDE FILE="inc_sha256.asp" -->
<!--#INCLUDE FILE="inc_header_short.asp" -->
<!--#INCLUDE FILE="inc_func_member.asp" -->
<% 
if Request.QueryString("ID") <> "" and IsNumeric(Request.QueryString("ID")) = True then
	intMemberID = cLng(Request.QueryString("ID"))
else
	intMemberID = 0
end if

'## Forum_SQL
strSql = "SELECT M.M_RECEIVE_EMAIL, M.M_EMAIL, M.M_NAME FROM " & strMemberTablePrefix & "MEMBERS M"
strSql = strSql & " WHERE M.MEMBER_ID = " & intMemberID

set rs = my_Conn.Execute (strSql)

Response.Write	"    <p><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>Send an E-MAIL Message</font></p>" & vbNewLine

if rs.bof or rs.eof then
	Response.Write	"    <p><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHiLiteFontColor & """>There is no Member with that Member ID</font></p>" & vbNewLine
else
	strRName = ChkString(rs("M_NAME"),"display")
	strREmail = rs("M_EMAIL")
	strRReceiveEmail = rs("M_RECEIVE_EMAIL")
	
	rs.close
	set rs = nothing
	
	if mLev > 2 or strRReceiveEmail = "1" then
		if lcase(strEmail) = "1" then
			if Request.QueryString("mode") = "DoIt" then
				Err_Msg = ""
				if mLev => 2 then
					strSql =  "SELECT M_NAME, M_USERNAME, M_EMAIL "
					strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS"
					strSql = strSql & " WHERE MEMBER_ID = " & MemberID

					set rs2 = my_conn.Execute (strSql)
					YName = rs2("M_NAME")
					YEmail = rs2("M_EMAIL")
					set rs2 = nothing
				else
					YName = Request.Form("YName")
					YEmail = Request.Form("YEmail")
					if YName = "" then
						Err_Msg = Err_Msg & "<li>You must enter your UserName</li>"
					end if
					if YEmail = "" then 
						Err_Msg = Err_Msg & "<li>You must give your e-mail address</li>"
					else
						if EmailField(YEmail) = 0 then 
							Err_Msg = Err_Msg & "<li>You must enter a valid e-mail address</li>"
						end if
					end if
				end if
				if Request.Form("Msg") = "" then 
					Err_Msg = Err_Msg & "<li>You must enter a message</li>"
				end if
				'##  E-mails Message to the Author of this Reply.  
				if (Err_Msg = "") then
					strRecipientsName = strRName
					strRecipients = strREmail
					strFrom = YEmail
					strFromName = YName
					strSubject = "Sent From " & strForumTitle & " by " & YName
					strMessage = "Hello " & strRName & vbNewline & vbNewline
					strMessage = strMessage & "You received the following message from: " & YName & " (" & YEmail & ") " & vbNewline & vbNewline 
					strMessage = strMessage & "At: " & strForumURL & vbNewline & vbNewline
					strMessage = strMessage & Request.Form("Msg") & vbNewline & vbNewline

					if strFrom <> "" then 
						strSender = strFrom
					end if
%>
<!--#INCLUDE FILE="inc_mail.asp" -->
<%
					Response.Write	"    <p><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>E-mail has been sent</font></p>" & vbNewLine
				else
					Response.Write	"    <p><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """>There Was A Problem With Your E-mail</font></p>" & vbNewLine
					Response.Write	"      <table>" & vbNewLine & _
							"        <tr>" & vbNewLine & _
							"          <td><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHiLiteFontColor & """><ul>" & Err_Msg & "</ul></font></td>" & vbNewLine & _
							"        </tr>" & vbNewLine & _
							"      </table>" & vbNewLine & _
							"    <p><font size=""" & strDefaultFontSize & """><a href=""JavaScript:history.go(-1)"">Go Back To Enter Data</a></font></p>" & vbNewLine
					WriteFooterShort
					Response.End 
				end if
			else 
				Err_Msg = ""
				if trim(strREmail) <> "" then
					strSql =  "SELECT M_NAME, M_USERNAME, M_EMAIL "
					strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS"
					strSql = strSql & " WHERE MEMBER_ID = " & MemberID

					set rs2 = my_conn.Execute (strSql)
					YName = ""
					YEmail = ""

					if (rs2.EOF or rs2.BOF)  then
						if strLogonForMail <> "0" then 
							Err_Msg = Err_Msg & "<li>You must be logged on to send a message</li>"

							Response.Write	"      <table>" & vbNewLine & _
									"        <tr>" & vbNewLine & _
									"          <td><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHiLiteFontColor & """><ul>" & Err_Msg & "</ul></font></td>" & vbNewLine & _
									"        </tr>" & vbNewLine & _
									"      </table>" & vbNewLine
							WriteFooterShort
							Response.End
						end if
					else
						YName = Trim("" & rs2("M_NAME"))
						YEmail = Trim("" & rs2("M_EMAIL"))
					end if
					rs2.close
					set rs2 = nothing

					Response.Write	"      <form action=""pop_mail.asp?mode=DoIt&id=" & intMemberID & """ method=""Post"" id=""Form1"" name=""Form1"">" & vbNewLine & _
							"      <table border=""0"" width=""90%"" cellspacing=""0"" cellpadding=""0"">" & vbNewLine & _
							"        <tr>" & vbNewLine & _
							"          <td bgcolor=""" & strPopUpBorderColor & """>" & vbNewLine & _
							"            <table border=""0"" width=""100%"" cellspacing=""1"" cellpadding=""1"">" & vbNewLine & _
							"              <tr>" & vbNewLine & _
							"                <td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Send To Name:</font></b></td>" & vbNewLine & _
							"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & strRName & "</font></td>" & vbNewLine & _
							"              </tr>" & vbNewLine & _
							"              <tr>" & vbNewLine & _
							"                <td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Your Name:</font></b></td>" & vbNewLine & _
							"                <td bgColor=""" & strPopUpTableColor & """>"
					if YName = "" then
						Response.Write "<input name=""YName"" type=""text"" value=""" & YName & """ size=""25"">"
					else
						Response.Write "<font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & YName & "</font>" & vbNewLine
					end if
					Response.Write	"              </td></tr>" & vbNewLine & _
							"              <tr>" & vbNewLine & _
							"                <td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Your E-mail:</font></b></td>" & vbNewLine & _
							"                <td bgColor=""" & strPopUpTableColor & """>"
					if YEmail = "" then
						Response.Write "<input name=""YEmail"" type=""text"" value=""" & YEmail & """ size=""25"">"
					else
						Response.Write "<font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & YEmail & "</font>"
					end if
					Response.Write	"</td>" & vbNewLine & _
							"              </tr>" & vbNewLine & _
							"              <tr>" & vbNewLine & _
							"                <td bgColor=""" & strPopUpTableColor & """ colspan=""2""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Message:</font></b></td>" & vbNewLine & _
							"              </tr>" & vbNewLine & _
							"              <tr>" & vbNewLine & _
							"                <td bgColor=""" & strPopUpTableColor & """ colspan=""2""><textarea name=""Msg"" cols=""40"" rows=""5""></textarea></td>" & vbNewLine & _
							"              </tr>" & vbNewLine & _
							"              <tr>" & vbNewLine & _
							"                <td bgColor=""" & strPopUpTableColor & """ colspan=""2"" align=""center""><input type=""Submit"" value=""Send"" id=""Submit1"" name=""Submit1""></td>" & vbNewLine & _
							"              </tr>" & vbNewLine & _
							"            </table>" & vbNewLine & _
							"          </td>" & vbNewLine & _
							"        </tr>" & vbNewLine & _
							"      </table>" & vbNewLine & _
							"      </form>" & vbNewLine
				else
					Response.Write	"    <p><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHiLiteFontColor & """>No E-mail address is available for this user.</font></p>" & vbNewLine
				end if
			end if
		else
			Response.Write	"    <p><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Click to send <a href=""mailto:" & chkString(rs("M_EMAIL"),"display") & """>" & strRName & "</a> an e-mail</font></p>" & vbNewLine
		end if
	else
		Response.Write	"    <p><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHiLiteFontColor & """>This Member does not wish to receive e-mail.</font></p>" & vbNewLine
	end if
end if
WriteFooterShort
Response.End
%>
