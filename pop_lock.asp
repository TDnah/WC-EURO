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
<!--#INCLUDE FILE="inc_sha256.asp"-->
<!--#INCLUDE FILE="inc_header_short.asp" -->
<%
if Request("CAT_ID") <> "" then
	if IsNumeric(Request("CAT_ID")) = True then Cat_ID = cLng(Request("CAT_ID")) else Cat_ID = 0
end if
if Request("FORUM_ID") <> "" then
	if IsNumeric(Request("FORUM_ID")) = True then Forum_ID = cLng(Request("FORUM_ID")) else Forum_ID = 0
end if
if Request("TOPIC_ID") <> "" then
	if IsNumeric(Request("TOPIC_ID")) = True then Topic_ID = cLng(Request("TOPIC_ID")) else Topic_ID = 0
end if
if Request("REPLY_ID") <> "" then
	if IsNumeric(Request("REPLY_ID")) = True then Reply_ID = cLng(Request("REPLY_ID")) else Reply_ID = 0
end if
if Request("MEMBER_ID") <> "" then
	if IsNumeric(Request("MEMBER_ID")) = True then Member_ID = cLng(Request("MEMBER_ID")) else Member_ID = 0
end if

if (Cat_ID + Forum_ID + Topic_ID + Reply_ID + Member_ID) < 1 then
	Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """><b>The URL has been modified!</b></font></p>" & vbNewLine & _
			"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHiLiteFontColor & """><b>Possible Hacking Attempt!</b></font></p>" & vbNewLine
	WriteFooterShort
	Response.End
end if

Mode_Type = ChkString(Request("mode"), "SQLString")
strPassword = trim(Request.Form("pass"))

select case Mode_Type 
	case "CloseTopic"
		strEncodedPassword = sha256("" & strPassword)
		mLev = cLng(chkUser(strDBNTFUserName, strEncodedPassword,-1))
		if mLev > 0 then  '## is Member
			if (chkForumModerator(Forum_ID, strDBNTFUserName) = "1") or (mLev = 4) then
				'## Forum_SQL
				strSql = "UPDATE " & strTablePrefix & "TOPICS "
				strSql = strSql & " SET T_STATUS = " & 0
				if Request.Form("noArchiveFlag") = "1" then
					strSQL = strSql & ", T_ARCHIVE_FLAG = " & 0
				else
					strSQL = strSql & ", T_ARCHIVE_FLAG = " & 1
				end if
				strSql = strSql & " WHERE TOPIC_ID = " & Topic_ID 

				my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

				Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """><b>Topic Locked!</b></font></p>" & vbNewLine & _
						"      <script language=""javascript1.2"">self.opener.location.reload();</script>" & vbNewLine
			else
				Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """><b>No Permissions to Lock Topic</b></font><br />" & _
						"<br /><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""JavaScript:onClick=history.go(-1)"">Go Back to Re-Authenticate</a></font></p>" & vbNewLine
			end if
		else
			Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """><b>No Permissions to Lock Topic</b></font><br />" & _
					"<br /><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""JavaScript:onClick=history.go(-1)"">Go Back to Re-Authenticate</a></font></p>" & vbNewLine
		end if 

	case "CloseForum"

		strEncodedPassword = sha256("" & strPassword)
		mLev = cLng(chkUser(strDBNTFUserName, strEncodedPassword,-1))
		if mLev > 0 then  '## is Member
			if (chkForumModerator(Forum_ID, strDBNTFUserName) = "1") or (mLev = 4) then
				'## Forum_SQL
				strSql = "UPDATE " & strTablePrefix & "FORUM "
				strSql = strSql & " SET F_STATUS = 0 "
				strSql = strSql & " WHERE FORUM_ID = " & Forum_ID 

				my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

				Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """><b>Forum Locked!</b></font></p>" & vbNewLine & _
						"      <script language=""javascript1.2"">self.opener.location.reload();</script>" & vbNewLine
			else
				Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """><b>No Permissions to Lock Forum</b></font><br />" & _
						"<br /><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""JavaScript:onClick=history.go(-1)"">Go Back to Re-Authenticate</a></font></p>" & vbNewLine
			end if
		else
			Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """><b>No Permissions to Lock Forum</b></font><br />" & _
					"<br /><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""JavaScript:onClick=history.go(-1)"">Go Back to Re-Authenticate</a></font></p>" & vbNewLine
		end if 

	case "CloseCategory"

		strEncodedPassword = sha256("" & strPassword)
		mLev = cLng(ChkUser(strDBNTFUserName, strEncodedPassword,-1)) 
		if mLev > 0 then  '## is Member
			if mLev = 4 then
				'## Forum_SQL
				strSql = "UPDATE " & strTablePrefix & "CATEGORY "
				strSql = strSql & " SET CAT_STATUS = 0 "
				strSql = strSql & " WHERE CAT_ID = " & Cat_ID 

				my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

				Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """><b>Category Locked!</b></font></p>" & vbNewLine & _
						"      <script language=""javascript1.2"">self.opener.location.reload();</script>" & vbNewLine
			else
				Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """><b>No Permissions to Lock Category</b></font><br />" & _
						"<br /><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""JavaScript:onClick=history.go(-1)"">Go Back to Re-Authenticate</a></font></p>" & vbNewLine
			end if
		else
			Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """><b>No Permissions to Lock Category</b></font><br />" & _
					"<br /><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""JavaScript:onClick=history.go(-1)"">Go Back to Re-Authenticate</a></font></p>" & vbNewLine
		end if 
	case "LockMember"
		strEncodedPassword = sha256("" & strPassword)
		mLev = cLng(ChkUser(strDBNTFUserName, strEncodedPassword,-1)) 
		if mLev > 0 then  '## is Member
			if (mLev = 4) and (cLng(chkCanLock(MemberID,Member_ID)) = 1) then
				'## Forum_SQL
				strSql = "UPDATE " & strMemberTablePrefix & "MEMBERS "
				strSql = strSql & " SET M_STATUS = 0 "
				strSql = strSql & " WHERE MEMBER_ID = " & Member_ID 

				my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

				Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """><b>Member Locked!</b></font></p>" & vbNewLine & _
						"      <script language=""javascript1.2"">self.opener.location.reload();</script>" & vbNewLine
			else
				Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """><b>No Permissions to Lock a Member</b></font><br />" & _
						"<br /><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""JavaScript:onClick=history.go(-1)"">Go Back to Re-Authenticate</a></font></p>" & vbNewLine
			end if
		else
			Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """><b>No Permissions to Lock a Member</b></font><br />" & _
					"<br /><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""JavaScript:onClick=history.go(-1)"">Go Back to Re-Authenticate</a></font></p>" & vbNewLine
		end if 
	case "StickyTopic"
		strEncodedPassword = sha256("" & strPassword)
		mLev = cLng(chkUser(strDBNTFUserName, strEncodedPassword,-1))
		if mLev > 0 then  '## is Member
			if (chkForumModerator(Forum_ID, strDBNTFUserName) = "1") or (mLev = 4) then
				'## Forum_SQL
				strSql = "UPDATE " & strTablePrefix & "TOPICS "
				strSql = strSql & " SET T_STICKY = " & 1
				strSQL = strSql & ", T_ARCHIVE_FLAG = " & 0
				strSql = strSql & " WHERE TOPIC_ID = " & Topic_ID 

				my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

				Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """><b>Topic Made Sticky!</b></font></p>" & vbNewLine & _
						"      <script language=""javascript1.2"">self.opener.location.reload();</script>" & vbNewLine
			else
				Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """><b>No Permissions to Make Topic Sticky!</b></font><br />" & _
						"<br /><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""JavaScript:onClick=history.go(-1)"">Go Back to Re-Authenticate</a></font></p>" & vbNewLine
			end if
		else
			Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """><b>No Permissions to Make Topic Sticky!</b></font><br />" & _
					"<br /><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""JavaScript:onClick=history.go(-1)"">Go Back to Re-Authenticate</a></font></p>" & vbNewLine
		end if 
	case else 
		Response.Write	"      <p><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>"
		select case Mode_Type
			case "Topic"
				Response.Write("Lock Topic")
			case "Forum"
				Response.Write("Lock Forum")
			case "Category"
				Response.Write("Lock Category")
			case "Member"
				Response.Write("Lock Member")
			case "STopic"
				Response.Write("Make Topic Sticky")
		end select
		Response.Write	"</font></p>" & vbNewLine & _
				"      <p><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b><font  color=""" & strHiLiteFontColor & """>NOTE:&nbsp;</font></b>"
		select case Mode_Type
			case "STopic"
				Response.Write("Only Moderators and Administrators can make a Topic Sticky.")
			case "Member"
				Response.Write("Only Administrators can lock a Member.")
			case "Category"
				Response.Write("Only Administrators can lock a Category.")
			case "Forum"
				Response.Write("Only Moderators and Administrators can lock a Forum.")
			case "Topic"
				Response.Write("Only Moderators and Administrators can lock a Topic.")
		end select
		Response.Write	"</font></p>" & vbNewLine & _
				"      <form action=""pop_lock.asp?mode="
		select case Mode_Type
			case "Topic"
				Response.Write("CloseTopic")
			case "Forum"
				Response.Write("CloseForum")
			case "Category"
				Response.Write("CloseCategory")
			case "Member"
				Response.Write("LockMember")
			case "STopic"
				Response.Write("StickyTopic")
		end select
		Response.Write	""" method=""post"" id=""Form1"" name=""Form1"">" & vbNewLine & _
				"      <input type=""hidden"" name=""Method_Type"" value="""
		select case Mode_Type
			case "Topic"
				Response.Write("CloseTopic")
			case "Forum"
				Response.Write("CloseForum")
			case "Category"
				Response.Write("CloseCategory")
			case "Member"
				Response.Write("LockMember")
			case "STopic"
				Response.Write("StickyTopic")
		end select
		Response.Write	""">" & vbNewLine & _
				"      <input type=""hidden"" name=""TOPIC_ID"" value=""" & Topic_ID & """>" & vbNewLine & _
				"      <input type=""hidden"" name=""FORUM_ID"" value=""" & Forum_ID & """>" & vbNewLine & _
				"      <input type=""hidden"" name=""CAT_ID"" value=""" & Cat_ID & """>" & vbNewLine & _
				"      <input type=""hidden"" name=""MEMBER_ID"" value=""" & Member_ID & """>" & vbNewLine & _
				"      <table border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbNewLine & _
				"        <tr>" & vbNewLine & _
				"          <td bgcolor=""" & strPopUpBorderColor & """>" & vbNewLine & _
				"            <table border=""0"" width=""100%"" cellspacing=""1"" cellpadding=""1"">" & vbNewLine
		if strAuthType="db" then
			Response.Write	"              <tr>" & vbNewLine & _
					"                <td bgcolor=""" & strPopUpTableColor & """ align=""right"" nowrap><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>User Name:</font></b></td>" & vbNewLine & _
					"                <td bgcolor=""" & strPopUpTableColor & """><input type=""text"" name=""Name"" maxLength=""25"" value=""" & chkString(strDBNTUserName,"display") & """ style=""width:150px;""></td>" & vbNewLine & _
					"              </tr>" & vbNewLine & _
					"              <tr>" & vbNewLine & _
					"                <td bgcolor=""" & strPopUpTableColor & """ align=""right"" nowrap><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Password:</font></b></td>" & vbNewLine & _
					"                <td bgcolor=""" & strPopUpTableColor & """><input type=""Password"" name=""Pass"" maxLength=""25"" value="""" style=""width:150px;""></td>" & vbNewLine & _
					"              </tr>" & vbNewLine
		else
			if strAuthType="nt" then
				Response.Write	"              <tr>" & vbNewLine & _
						"                <td bgcolor=""" & strPopUpTableColor & """ align=""right"" nowrap><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>NT Account:</font></b></td>" & vbNewLine & _
						"                <td bgcolor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & chkString(strDBNTUserName,"display") & "</font></td>" & vbNewLine & _
						"              </tr>" & vbNewLine
			end if
		end if 
		if Mode_Type = "Topic" Then  
			response.write	"              <tr>" & vbNewLine
			strSQL = "SELECT T_ARCHIVE_FLAG FROM " & strTablePrefix & "TOPICS "
			strSql = strSQL & "WHERE TOPIC_ID = " & Topic_ID 
			set rs = my_conn.Execute(strSql)
			response.write	"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Do not Archive</font></b></td>" & vbNewLine & _
					"                <td bgColor=""" & strPopUpTableColor & """ align=""left""><input type=""Checkbox"" value=""1"" name=""noArchiveFlag"""
			if rs("T_ARCHIVE_FLAG") = 0 then response.write(" checked")
			response.write	"></td>" & vbNewLine & _
					"              </tr>" & vbNewLine
	 		rs.close
		end If 
		Response.Write	"              <tr>" & vbNewLine & _
				"                <td bgColor=""" & strPopUpTableColor & """ colspan=""2"" align=""center""><Input type=""Submit"" value=""Send""></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"            </table>" & vbNewLine & _
				"          </td>" & vbNewLine & _
				"        </tr>" & vbNewLine & _
				"      </table>" & vbNewLine & _
				"      </form>" & vbNewLine
end select
WriteFooterShort
Response.End

function chkCanLock(fAM_ID, fM_ID)
	'## Forum_SQL
	strSql = "SELECT MEMBER_ID, M_LEVEL "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
	StrSql = strSql & " WHERE MEMBER_ID = " & fM_ID

	set rsCheck = my_Conn.Execute (strSql)

	if rsCheck.BOF or rsCheck.EOF then
		chkCanLock = 0 '## No Members Found
	else
		if cLng(rsCheck("MEMBER_ID")) = cLng(fAM_ID) then 
			chkCanLock = 0 '## Can't lock self
		else
			Select case cLng(rsCheck("M_LEVEL"))
				case 1
					chkCanLock = 1 '## Can lock Normal User
				case 2
					chkCanLock = 1 '## Can lock Moderator
				case 3
					if fAM_ID <> intAdminMemberID then
						chkCanLock = 0 '## Only the Forum Admin can lock other Administrators
					else
						chkCanLock = 1 '## Forum Admin is ok to lock other Administrators
					end if
				case else
					chkCanLock = 0 '## Member doesn't have a Member Level?
			End Select
		end if	
	end if

	rsCheck.close	
	set rsCheck = nothing
end function
%>
