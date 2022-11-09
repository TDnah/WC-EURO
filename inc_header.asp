<script language="javascript" type="text/javascript">
    //var msg = "<ul><li>Tình hình tài chính sau buổi sơ kết <a href='VCWC2010.soket.xls'>xem tại đây</a></li><li>Giai đoạn 2 chỉ có 1 phần quà duy nhất dành cho nhà tài trợ kim cương là 1 chiếc kèn Vuvuzela. Nếu có nhiều nhà đồng tài trợ thì phần quà sẽ được chia đều</li><li>Điều kiện cần để xét danh hiệu nhà tài trợ kim cương là PHẢI bắt tất cả các trận</li></ul>";
    //var msg = "<h2 align=center>THÔNG BÁO</h2><ul><li>Buổi tổng kết chương trình Vui cùng Worldcup 2010 sẽ được tổ chức vào lúc 18h00 ngày 15/07/2010 tại nhà hàng Vườn Phố, số A2 Phan Đình Giót, P.2, Q.TB (Trong khuôn viên SVĐ QK7).</li><li>BTC sẽ thu tiền của các nhà tài trợ vào ngày 13/07/2010 </li></ul>&nbsp;&nbsp;&nbsp;BTC VCWC2010"
    function myPop() {
        this.square = null;
        this.overdiv = null;

        this.popOut = function(msgtxt) {
            //filter:alpha(opacity=25);-moz-opacity:.25;opacity:.25;
            this.overdiv = document.createElement("div");
            this.overdiv.className = "overdiv";

            this.square = document.createElement("div");
            this.square.className = "square";
            this.square.Code = this;
            var msg = document.createElement("div");
            msg.className = "msg";
            msg.innerHTML = msgtxt;
            this.square.appendChild(msg);
            var closebtn = document.createElement("button");
            closebtn.onclick = function() {
                this.parentNode.Code.popIn();
            }
            closebtn.innerHTML = " >> Đã đọc >> ";
            this.square.appendChild(closebtn);

            document.body.appendChild(this.overdiv);
            document.body.appendChild(this.square);
        }
        this.popIn = function() {
            if (this.square != null) {
                document.body.removeChild(this.square);
                this.square = null;
            }
            if (this.overdiv != null) {
                document.body.removeChild(this.overdiv);
                this.overdiv = null;
            }

        }
    }

    if (msg != "") {
        var pop = new myPop();
        window.onload = pop.popOut(msg);
    }

</script>

<style type="text/css">
 div.overdiv { filter: alpha(opacity=75);
                      -moz-opacity: .75;
                      opacity: .75;
                      position: absolute;
                      top: 0px;
                      left: 0px;
                      width: 100%; height: 100%; 
                      z-index:-1;}

        div.square { position: absolute;
                     top: 300px;
                     left: 300px;
                     padding:10px;
                     background-color:Red;
                     border: #f9f9f9;
                     height: 300px;
                     width: 500px; 
                     text-align:center;
                     }
        div.square div.msg { color: white;
                             font-size: 18px;
                             
                             text-align:left;}
</style>




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
Session.CodePage = 65001
%>
<!--#INCLUDE FILE="inc_func_common.asp" -->
<%


if strShowTimer = "1" then
	'### start of timer code
	Dim StopWatch(19) 

	sub StartTimer(x)
		StopWatch(x) = timer
	end sub

	function StopTimer(x)
		EndTime = Timer

		'Watch for the midnight wraparound...
		if EndTime < StopWatch(x) then
			EndTime = EndTime + (86400)
		end if

		StopTimer = EndTime - StopWatch(x)
	end function

	StartTimer 1

	'### end of timer code
end if

strArchiveTablePrefix = strTablePrefix & "A_"
strScriptName = request.servervariables("script_name")
strReferer = chkString(request.servervariables("HTTP_REFERER"),"refer")

if Application(strCookieURL & "down") then 
	if not Instr(strScriptName,"admin_") > 0 then
		Response.redirect("down.asp")
	end if
end if

if strPageBGImageURL = "" then
	strTmpPageBGImageURL = ""
elseif Instr(strPageBGImageURL,"/") > 0 or Instr(strPageBGImageURL,"\") > 0 then
	strTmpPageBGImageURL = " background=""" & strPageBGImageURL & """"
else
	strTmpPageBGImageURL = " background=""" & strImageUrl & strPageBGImageURL & """"
end if

If strDBType = "" then 
	Response.Write	"<html>" & vbNewLine & _
			"<head>" & vbNewline & _
			"<META http-equiv=Content-Type content=""" & "text/html; charset=utf-8""" & ">" & vbNewline & _
			"<title>" & strForumTitle & "</title>" & vbNewline


'## START - REMOVAL, MODIFICATION OR CIRCUMVENTING THIS CODE WILL VIOLATE THE SNITZ FORUMS 2000 LICENSE AGREEMENT
Response.Write	"<meta name=""copyright"" content=""This Forum code is Copyright (C) 2000-06 Michael Anderson, Pierre Gorissen, Huw Reddick and Richard Kinser, Non-Forum Related code is Copyright (C) " & strCopyright & """>" & vbNewline 
'## END   - REMOVAL, MODIFICATION OR CIRCUMVENTING THIS CODE WILL VIOLATE THE SNITZ FORUMS 2000 LICENSE AGREEMENT

	Response.Write	"</head>" & vbNewLine & _
			"<body" & strTmpPageBGImageURL & " bgColor=""" & strPageBGColor & """ text=""" & strDefaultFontColor & """ link=""" & strLinkColor & """ aLink=""" & strActiveLinkColor & """ vLink=""" & strVisitedLinkColor & """>" & vbNewLine & _
			"<table border=""0"" cellspacing=""0"" cellpadding=""5"" width=""50%"" height=""40%"" align=""center"">" & vbNewLine & _
			"  <tr>" & vbNewLine & _
			"    <td bgColor=""navyblue"" align=""center""><p><font face=""Verdana, Arial, Helvetica"" size=""2"">" & _
			"<b>There has been a problem...</b><br /><br />" & _
			"Your <b>strDBType</b> is not set, please edit your <b>config.asp</b><br />to reflect your database type." & _
			"</font></p></td>" & vbNewLine & _
			"  </tr>" & vbNewLine & _
			"  <tr>" & vbNewLine & _
			"    <td align=""center""><font face=""Verdana, Arial, Helvetica"" size=""2"">" & _
			"<a href=""default.asp"" target=""_top"">Click here to retry.</a></font></td>" & vbNewLine & _
			"  </tr>" & vbNewLine & _
			"</table>" & vbNewLine & _
			"</body>" & vbNewLine & _
			"</html>" & vbNewLine
	Response.End
end if

set my_Conn = Server.CreateObject("ADODB.Connection")
my_Conn.Open strConnString

if (strAuthType = "nt") then
	call NTauthenticate()
	if (ChkAccountReg() = "1") then
		call NTUser()
	end if
end if

if strGroupCategories = "1" then
	if Request.QueryString("Group") = "" then
		if Request.Cookies(strCookieURL & "GROUP") = "" Then
			Group = 2
		else 
			Group = cLng(Request.Cookies(strCookieURL & "GROUP"))
		end if
	else
		Group = cLng(Request.QueryString("Group"))
	end if
	'set default
	Session(strCookieURL & "GROUP_ICON") = "icon_group_categories.gif"
	Session(strCookieURL & "GROUP_IMAGE") = strTitleImage
	'Forum_SQL - Group exists ?
	strSql = "SELECT GROUP_ID, GROUP_NAME, GROUP_ICON, GROUP_IMAGE " 
	strSql = strSql & " FROM " & strTablePrefix & "GROUP_NAMES "
	strSql = strSql & " WHERE GROUP_ID = " & Group
	set rs2 = my_Conn.Execute (strSql)
	if rs2.EOF or rs2.BOF then
		Group = 2
		strSql = "SELECT GROUP_ID, GROUP_NAME, GROUP_ICON, GROUP_IMAGE " 
		strSql = strSql & " FROM " & strTablePrefix & "GROUP_NAMES "
		strSql = strSql & " WHERE GROUP_ID = " & Group
		set rs2 = my_Conn.Execute (strSql)
	end if	
	Session(strCookieURL & "GROUP_NAME") = rs2("GROUP_NAME")
	if instr(rs2("GROUP_ICON"), ".") then
		Session(strCookieURL & "GROUP_ICON") = rs2("GROUP_ICON")
	end if
	if instr(rs2("GROUP_IMAGE"), ".") then
		Session(strCookieURL & "GROUP_IMAGE") = rs2("GROUP_IMAGE")
	end if
	rs2.Close  
	set rs2 = nothing  
	Response.Cookies(strCookieURL & "GROUP") = Group
	Response.Cookies(strCookieURL & "GROUP").Expires =  dateAdd("d", intCookieDuration, strForumTimeAdjust)
	if Session(strCookieURL & "GROUP_IMAGE") <> "" then
		strTitleImage = Session(strCookieURL & "GROUP_IMAGE") 
	end if 
end if

strDBNTUserName = Request.Cookies(strUniqueID & "User")("Name")
strDBNTFUserName = trim(chkString(Request.Form("Name"),"SQLString"))
if strDBNTFUserName = "" then strDBNTFUserName = trim(chkString(Request.Form("User"),"SQLString"))
if strAuthType = "nt" then
	strDBNTUserName = Session(strCookieURL & "userID")
	strDBNTFUserName = Session(strCookieURL & "userID")
end if

if strRequireReg = "1" and strDBNTUserName = "" then
	if not Instr(strScriptName,"register.asp") > 0 and _
	not Instr(strScriptName,"password.asp") > 0 and _
	not Instr(strScriptName,"faq.asp") > 0 and _
	not Instr(strScriptName,"login.asp") > 0 then
		scriptname = split(request.servervariables("SCRIPT_NAME"),"/")
		if Request.QueryString <> "" then
			Response.Redirect("login.asp?target=" & lcase(scriptname(ubound(scriptname))) & "?" & Request.QueryString)
		else
			Response.Redirect("login.asp?target=" & lcase(scriptname(ubound(scriptname))))
		end if
	end if
end if

select case Request.Form("Method_Type")
	case "login"
		strEncodedPassword = sha256("" & Request.Form("Password"))
		select case chkUser(strDBNTFUserName, strEncodedPassword,-1)
			case 1, 2, 3, 4
				Call DoCookies(Request.Form("SavePassword"))
				strLoginStatus = 1
			case else
				strLoginStatus = 0
		end select
	case "logout"
		Call ClearCookies()
end select

'Dung Add for USER ONLINE
if Request.Form("Method_Type")<>"logout" then
		session("UserID")=strDBNTUserName & ""
		strSql ="UPDATE " & strMemberTablePrefix & "MEMBERS "
		strSql = strSql & " SET IsOnline = -1, M_LASTHEREDATE = '" & DateToStr(now) & "'"
		strSql = strSql & " WHERE M_NAME = '" & strDBNTUserName & "'"

		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
End if
'End Dung Add for USER ONLINE

if trim(strDBNTUserName) <> "" and trim(Request.Cookies(strUniqueID & "User")("Pword")) <> "" then
	chkCookie = 1
	mLev = cLng(chkUser(strDBNTUserName, Request.Cookies(strUniqueID & "User")("Pword"),-1))
	chkCookie = 0
else
	MemberID = -1
	mLev = 0
end if

if mLev = 4 and strEmailVal = "1" and strRestrictReg = "1" and strEmail = "1" then
	'## Forum_SQL - Get membercount from DB 
	strSql = "SELECT COUNT(MEMBER_ID) AS U_COUNT FROM " & strMemberTablePrefix & "MEMBERS_PENDING WHERE M_APPROVE = " & 0

	set rs = Server.CreateObject("ADODB.Recordset")
	rs.open strSql, my_Conn

	if not rs.EOF then
		User_Count = cLng(rs("U_COUNT"))
	else
		User_Count = 0
	end if

	rs.close
	set rs = nothing
end if

Response.Write	"<html>" & vbNewline & vbNewline & _
		"<head>" & vbNewline & _
		"<META http-equiv=Content-Type content=""" & "text/html; charset=utf-8""" & ">" & vbNewline & _
		"<title>" & GetNewTitle(strScriptName) & "</title>" & vbNewline


'## START - REMOVAL, MODIFICATION OR CIRCUMVENTING THIS CODE WILL VIOLATE THE SNITZ FORUMS 2000 LICENSE AGREEMENT
Response.Write	"<meta name=""copyright"" content=""This Forum code is Copyright (C) 2000-06 Michael Anderson, Pierre Gorissen, Huw Reddick and Richard Kinser, Non-Forum Related code is Copyright (C) " & strCopyright & """>" & vbNewline 
'## END   - REMOVAL, MODIFICATION OR CIRCUMVENTING THIS CODE WILL VIOLATE THE SNITZ FORUMS 2000 LICENSE AGREEMENT



Response.Write	"<script language=""JavaScript"" type=""text/javascript"">" & vbNewLine & _
		"<!-- hide from JavaScript-challenged browsers" & vbNewLine & _
		"function openWindow(url) {" & vbNewLine & _
		"	popupWin = window.open(url,'new_page','width=400,height=400')" & vbNewLine & _
		"}" & vbNewLine & _
		"function openWindow2(url) {" & vbNewLine & _
		"	popupWin = window.open(url,'new_page','width=400,height=450')" & vbNewLine & _
		"}" & vbNewLine & _
		"function openWindow3(url) {" & vbNewLine & _
		"	popupWin = window.open(url,'new_page','width=400,height=450,scrollbars=yes')" & vbNewLine & _
		"}" & vbNewLine & _
		"function openWindow4(url) {" & vbNewLine & _
		"	popupWin = window.open(url,'new_page','width=400,height=525')" & vbNewLine & _
		"}" & vbNewLine & _
		"function openWindow5(url) {" & vbNewLine & _
		"	popupWin = window.open(url,'new_page','width=450,height=525,scrollbars=yes,toolbars=yes,menubar=yes,resizable=yes')" & vbNewLine & _
		"}" & vbNewLine & _
		"function openWindow6(url) {" & vbNewLine & _
		"	popupWin = window.open(url,'new_page','width=500,height=450,scrollbars=yes')" & vbNewLine & _
		"}" & vbNewLine & _
		"function openWindowHelp(url) {" & vbNewLine & _
		"	popupWin = window.open(url,'new_page','width=470,height=200,scrollbars=yes')" & vbNewLine & _
		"}" & vbNewLine & _
		"// done hiding -->" & vbNewLine & _
		"</script>" & vbNewLine & _
		"<style type=""text/css"">" & vbNewLine & _
		"<!--" & vbNewLine & _
		"a:link    {color:" & strLinkColor & ";text-decoration:" & strLinkTextDecoration & "}" & vbNewLine & _
		"a:visited {color:" & strVisitedLinkColor & ";text-decoration:" & strVisitedTextDecoration & "}" & vbNewLine & _
		"a:hover   {color:" & strHoverFontColor & ";text-decoration:" & strHoverTextDecoration & "}" & vbNewLine & _
		"a:active  {color:" & strActiveLinkColor & ";text-decoration:" & strActiveTextDecoration & "}" & vbNewLine & _
		".spnMessageText a:link    {color:" & strForumLinkColor & ";text-decoration:" & strForumLinkTextDecoration & "}" & vbNewLine & _
		".spnMessageText a:visited {color:" & strForumVisitedLinkColor & ";text-decoration:" & strForumVisitedTextDecoration & "}" & vbNewLine & _
		".spnMessageText a:hover   {color:" & strForumHoverFontColor & ";text-decoration:" & strForumHoverTextDecoration & "}" & vbNewLine & _
		".spnMessageText a:active  {color:" & strForumActiveLinkColor & ";text-decoration:" & strForumActiveTextDecoration & "}" & vbNewLine & _
		".spnSearchHighlight {background-color:" & strSearchHiLiteColor & "}" & vbNewLine & _
		"input.radio {background:" & strPopUpTableColor & ";color:#000000}" & vbNewLine & _
		"-->" & vbNewLine & _
		"</style>" & vbNewLine & _
		"</head>" & vbNewLine & _
		vbNewLine & _
		"<body" & strTmpPageBGImageURL & " bgColor=""" & strPageBGColor & """ text=""" & strDefaultFontColor & """ link=""" & strLinkColor & """ aLink=""" & strActiveLinkColor & """ vLink=""" & strVisitedLinkColor & """>" & vbNewLine & _
		"<a name=""top""></a>" & vbNewLine & _
		vbNewLine & _
		"<table align=""center"" border=""0"" cellPadding=""0"" cellSpacing=""0"" width=""100%"">" & vbNewLine & _
		"  <tr>" & vbNewLine & _
		"    <td valign=""top"" width=""300px""><a href=""home.asp"" tabindex=""-1"">" & getCurrentIcon(strTitleImage & "||",strForumTitle,"") & "</a></td>" & vbNewLine & _
        "    <td align=""center"" valign=""top"" width=""*""><img height=100px src=""/fb02.gif""><img height=100px src=""/fb10.gif""><img height=100px src=""/fb09.gif""></td>" & vbNewLine & _
		"    <td align=""center"" valign=""top"" width=""*"">" & vbNewLine & _
		"      <table border=""0"" cellPadding=""2"" cellSpacing=""0"">" & vbNewLine & _
		"        <tr>" & vbNewLine & _
		"          <td align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """>" & vbNewLine
call sForumNavigation()
Response.Write	"</font></td>" & vbNewLine & _
		"        </tr>" & vbNewLine

select case Request.Form("Method_Type")

	case "login"
		Response.Write	"      </table>" & vbNewLine & _
				"    </td>" & vbNewLine & _
				"  </tr>" & vbNewLine & _
				"</table>" & vbNewLine
		if strLoginStatus = 0 then
			Response.Write	"<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>Your username and/or password were incorrect.</font></p>" & vbNewLine & _
					"<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>Please either try again or register for an account.</font></p>" & vbNewLine
		else
			Response.Write	"<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>You logged on successfully!</font></p>" & vbNewLine & _
					"<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>Thank you for your participation.</font></p>" & vbNewLine
		end if
		Response.Write	"<meta http-equiv=""Refresh"" content=""2; URL=" & strReferer & """>" & vbNewLine & _
				"<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""" & strReferer & """>Back To Forum</font></a></p>" & vbNewLine & _
				"<table align=""center"" border=""0"" cellPadding=""0"" cellSpacing=""0"" width=""95%"">" & vbNewLine & _
				"  <tr>" & vbNewLine & _
				"    <td>" & vbNewLine
		WriteFooter
		Response.End
	case "logout" 
		Response.Write	"      </table>" & vbNewLine & _
				"    </td>" & vbNewLine & _
				"  </tr>" & vbNewLine & _
				"</table>" & vbNewLine & _
				"<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>You logged out successfully!</font></p>" & vbNewLine & _
				"<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>Thank you for your participation.</font></p>" & vbNewLine & _
				"<meta http-equiv=""Refresh"" content=""2; URL=default.asp"">" & vbNewLine & _
				"<p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""default.asp"">Back To Forum</font></a></p>" & vbNewLine & _
				"<table align=""center"" border=""0"" cellPadding=""0"" cellSpacing=""0"" width=""95%"">" & vbNewLine & _
				"  <tr>" & vbNewLine & _
				"    <td>" & vbNewLine
		WriteFooter
		Response.End
end select

if (mlev = 0) then
	if not(Instr(Request.ServerVariables("Path_Info"), "register.asp") > 0) and _
	not(Instr(Request.ServerVariables("Path_Info"), "pop_profile.asp") > 0) and _
	not(Instr(Request.ServerVariables("Path_Info"), "search.asp") > 0) and _
	not(Instr(Request.ServerVariables("Path_Info"), "login.asp") > 0) and _
	not(Instr(Request.ServerVariables("Path_Info"), "password.asp") > 0) and _
	not(Instr(Request.ServerVariables("Path_Info"), "faq.asp") > 0) and _
	not(Instr(Request.ServerVariables("Path_Info"), "post.asp") > 0) then
		Response.Write	"        <form action=""" & Request.ServerVariables("URL") & """ method=""post"" id=""form1"" name=""form1"">" & vbNewLine & _
				"        <input type=""hidden"" name=""Method_Type"" value=""login"">" & vbNewLine & _
				"        <tr>" & vbNewLine & _
				"          <td align=""center"">" & vbNewLine & _
				"            <table>" & vbNewLine & _
				"              <tr>" & vbNewLine
		if (strAuthType = "db") then
			Response.Write	"                <td><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """><b>Username:</b></font><br />" & vbNewLine & _
					"                <input type=""text"" name=""Name"" size=""10"" maxLength=""25"" value=""""></td>" & vbNewLine & _
					"                <td><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """><b>Password:</b></font><br />" & vbNewLine & _
					"                <input type=""password"" name=""Password"" size=""10"" maxLength=""25"" value=""""></td>" & vbNewLine & _
					"                <td valign=""bottom"">" & vbNewLine
			if strGfxButtons = "1" then
				Response.Write	"                <input src=""" & strImageUrl & "button_login.gif"" type=""image"" border=""0"" value=""Login"" id=""submit1"" name=""Login"">" & vbNewLine
			else
				Response.Write	"                <input type=""submit"" value=""Login"" id=""submit1"" name=""submit1"">" & vbNewLine
			end if 
			Response.Write	"                </td>" & vbNewLine & _
					"              </tr>" & vbNewLine & _
					"              <tr>" & vbNewLine & _
					"                <td colspan=""3"" align=""left""><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """>" & vbNewLine & _
					"                <input type=""checkbox"" name=""SavePassWord"" value=""true"" tabindex=""-1"" CHECKED><b> Save Password</b></font></td>" & vbNewLine
		else
			if (strAuthType = "nt") then 
				Response.Write	"                <td><font face=""" & strDefaultFontFace & """ size=""1""  color=""" & strHiLiteFontColor & """>Please <a href=""register.asp"" tabindex=""-1"">register</a> to post in these Forums</font></td>" & vbNewLine
			end if
		end if 
		Response.Write	"              </tr>" & vbNewLine
		if (lcase(strEmail) = "1") then
			Response.Write	"              <tr>" & vbNewLine & _
					"                <td colspan=""3"" align=""left""><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """>" & vbNewLine & _
					"                <a href=""password.asp""" & dWStatus("Choose a new password if you have forgotten your current one...") & " tabindex=""-1"">Forgot your "
			if strAuthType = "nt" then Response.Write("Admin ")
			Response.Write	"Password?</a>" & vbNewLine
			if (lcase(strNoCookies) = "1") then
				Response.Write	"                |" & vbNewLine & _
						"                <a href=""admin_home.asp""" & dWStatus("Access the Forum Admin Functions...") & " tabindex=""-1"">Admin Options</a>" & vbNewLine
			end if
			Response.Write	"                <br /><br /></font></td>" & vbNewLine & _
					"              </tr>" & vbNewLine
		end if
		Response.Write	"            </table>" & vbNewLine & _
				"          </td>" & vbNewLine & _
				"        </tr>" & vbNewLine & _
				"        </form>" & vbNewLine
	end if
else
	Response.Write	"        <form action=""" & Request.ServerVariables("URL") & """ method=""post"" id=""form2"" name=""form2"">" & vbNewLine & _
			"        <input type=""hidden"" name=""Method_Type"" value=""logout"">" & vbNewLine & _
			"        <tr>" & vbNewLine & _
			"          <td align=""center"">" & vbNewLine & _
			"            <table>" & vbNewLine & _
			"              <tr>" & vbNewLine & _
			"                <td align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>Xin ch&agrave;o&nbsp;"
	if strAuthType="nt" then
		Response.Write	"<b>" & Session(strCookieURL & "username") & "&nbsp;(" & Session(strCookieURL & "userid") & ")</b></font></td>" & vbNewLine & _
				"                <td>&nbsp;"
	else 
		if strAuthType = "db" then 
			Response.Write	"<b>" & profileLink(ChkString(strDBNTUserName, "display"),MemberID) & "</b></font></td>" & vbNewLine & _
					"                <td>"
			if strGfxButtons = "1" then
				Response.Write	"<input src=""" & strImageUrl & "button_logout.gif"" type=""image"" border=""0"" value=""Logout"" id=""submit1"" name=""Logout"" tabindex=""-1"">"
			else
				Response.Write	"<input type=""submit"" value=""Logout"" id=""submit1"" name=""submit1"" tabindex=""-1"">"
			end if 
		end if 
	end if 
	Response.Write	"</td>" & vbNewLine 

		
    'Dung add Clock
    Response.Write "<td align=""Center"">" & vbNewLine & _
    "<object classid=""clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"" codebase=""http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=9,0,28,0"" width=""107"" height=""30"" title=""Bây giờ là:"">" & vbNewLine & _
    "<param name=""movie"" value=""Clock.swf"" />" & vbNewLine & _
    "<param name=""quality"" value=""high"" />" & vbNewLine & _
    "<param name=""wmode"" value=""opaque"" />" & vbNewLine & _
    "<embed src=""Clock.swf"" quality=""high"" wmode=""opaque"" pluginspage=""http://www.adobe.com/shockwave/download/download.cgi?P1_Prod_Version=ShockwaveFlash"" type=""application/x-shockwave-flash"" width=""107"" height=""30""></embed>" & vbNewLine & _
    "</object>"  & vbNewLine & _
    "</td>" & vbNewLine 
    'End Dunf add Clock
    
    Response.Write "        </tr>" & vbNewLine & _
	"            </table>" & vbNewLine & _
	"          </td>" & vbNewLine & _
	"     </tr>" & vbNewLine
   
    
    
	if (mlev = 4) or (lcase(strNoCookies) = "1") then
		Response.Write	"        <tr>" & vbNewLine & _
				"          <td align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """><a href=""admin_home.asp""" & dWStatus("Access the Forum Admin Functions...") & " tabindex=""-1"">Admin Options</a>"
		if mLev = 4 and (strEmailVal = "1" and strRestrictReg = "1" and strEmail = "1" and User_Count > 0) then Response.Write("&nbsp;|&nbsp;<a href=""admin_accounts_pending.asp""" & dWStatus("(" & User_Count & ") Member(s) awaiting approval") & " tabindex=""-1"">(" & User_Count & ") Member(s) awaiting approval</a>")
		Response.Write	"<br /><br /></font></td>" & vbNewLine & _
				"        </tr>" & vbNewLine
	end if
	Response.Write	"        </form>" & vbNewLine
end if
Response.Write	"      </table>" & vbNewLine & _
		"    </td>" & vbNewLine & _
		"  </tr>" & vbNewLine & _
		"</table>" & vbNewLine & _
		"<table align=""center"" border=""0"" cellPadding=""0"" cellSpacing=""0"" width=""95%"">" & vbNewLine
'########### GROUP Categories ########### %>
<!--#INCLUDE FILE="inc_groupjump_to.asp" -->
<% '######## GROUP Categories ##############
Response.Write	"  <tr>" & vbNewLine & _
		"    <td>" & vbNewLine

sub sForumNavigation()
	' DEM --> Added code to show the subscription line
	if strSubscription > 0 and strEmail = "1" then
		if mlev > 0 then
			strSql = "SELECT COUNT(*) AS MySubCount FROM " & strTablePrefix & "SUBSCRIPTIONS"
			strSql = strSql & " WHERE MEMBER_ID = " & MemberID
			set rsCount = my_Conn.Execute (strSql)
			if rsCount.BOF or rsCount.EOF then
				' No Subscriptions found, do nothing
				MySubCount = 0
				rsCount.Close
				set rsCount = nothing
			else
				MySubCount = rsCount("MySubCount")
				rsCount.Close
				set rsCount = nothing
			end if
			if mLev = 4 then
				strSql = "SELECT COUNT(*) AS SubCount FROM " & strTablePrefix & "SUBSCRIPTIONS"
				set rsCount = my_Conn.Execute (strSql)
				if rsCount.BOF or rsCount.EOF then
					' No Subscriptions found, do nothing
					SubCount = 0
					rsCount.Close
					set rsCount = nothing
				else
					SubCount = rsCount("SubCount")
					rsCount.Close
					set rsCount = nothing
				end if
			end if
		else
			SubCount = 0
			MySubCount = 0
		end if
	else
		SubCount = 0
		MySubCount = 0
	end if
	Response.Write	"<font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """><a href=""" & "home.asp" & """" & dWStatus("Homepage") & " tabindex=""-1""><acronym title=""Trang ch&#7911;"">Trang ch&#7911;</acronym></a>" & " | " & "<a href=""" & "default.asp" & """" & "><acronym title=""Di&#7877;n &#273;&agrave;n"">Di&#7877;n &#273;&agrave;n</acronym></a>" & vbNewline & _
			"          |" & vbNewline
	if strUseExtendedProfile then 
		Response.Write	"          <a href=""pop_profile.asp?mode=Edit""" & dWStatus("Edit your personal profile...") & " tabindex=""-1""><acronym title=""Ch&#7881;nh s&#7917;a th&ocirc;ng tin c&aacute; nh&acirc;n..."">Th&ocirc;ng tin c&aacute; nh&acirc;n</acronym></a>" & vbNewline
	else
		Response.Write	"          <a href=""javascript:openWindow3('pop_profile.asp?mode=Edit')""" & dWStatus("Ch&#7881;nh s&#7917;a th&ocirc;ng tin c&aacute; nh&acirc;n...") & " tabindex=""-1""><acronym title=""Ch&#7881;nh s&#7917;a th&ocirc;ng tin c&aacute; nh&acirc;n..."">Th&ocirc;ng tin c&aacute; nh&acirc;n</acronym></a>" & vbNewline
	end if 

    if (mlev = 0) then
	if strAutoLogon <> "1" then
		if (strProhibitNewMembers <> "1") then
			Response.Write	"          |" & vbNewline & _
					"          <a href=""register.asp""" & dWStatus("&#272;&#259;ng k&yacute; &#273;&#7875; post b&agrave;i...") & " tabindex=""-1""><acronym title=""&#272;&#259;ng k&yacute; &#273;&#7875; post b&agrave;i..."">&#272;&#259;ng k&yacute;</acronym></a>" & vbNewline
		end if
	end if
	end if
	
	Response.Write	"          |" & vbNewline & _
			"          <a href=""active.asp""" & dWStatus("Xem tin m&#7899;i nh&#7845;t...") & " tabindex=""-1""><acronym title=""Xem tin m&#7899;i nh&#7845;t..."">Tin n&oacute;ng</acronym></a>" & vbNewline 
	' DEM --> Start of code added to show subscriptions if they exist
	if (strSubscription > 0) then
		if mlev = 4 and SubCount > 0 then
			Response.Write	"          |" & vbNewline & _
					"          <a href=""subscription_list.asp?MODE=all""" & dWStatus("See all current subscriptions") & " tabindex=""-1""><acronym title=""See all current subscriptions"">All Subscriptions</acronym></a>" & vbNewline
		end if
		if MySubCount > 0 then
			Response.Write	"          |" & vbNewline & _
					"          <a href=""subscription_list.asp""" & dWStatus("See all of your subscriptions") & " tabindex=""-1""><acronym title=""See all of your subscriptions"">My Subscriptions</acronym></a>" & vbNewline
		end if
	end if
	' DEM --> End of Code added to show subscriptions if they exist
	Response.Write	"          |" & vbNewline & _
			"          <a href=""members.asp""" & dWStatus("Danh s&aacute;ch th&agrave;nh vi&ecirc;n") & " tabindex=""-1""><acronym title=""Danh s&aacute;ch th&agrave;nh vi&ecirc;n"">Th&agrave;nh vi&ecirc;n</acronym></a>" & vbNewline & _
			"          |" & vbNewline & _
			"          <a href=""search.asp"
	if Request.QueryString("FORUM_ID") <> "" then Response.Write("?FORUM_ID=" & cLng(Request.QueryString("FORUM_ID")))
	Response.Write	"""" & dWStatus("T&igrave;m ki&#7871;m, ch&#432;a ch&#7855;c t&igrave;m th&#7845;y...") & " tabindex=""-1""><acronym title=""T&igrave;m ki&#7871;m, ch&#432;a ch&#7855;c t&igrave;m th&#7845;y..."">T&igrave;m ki&#7871;m</acronym></a>" & vbNewline & _
			"          |" & vbNewline & _
			"          <a href=""faq.asp""" & dWStatus("H&#7887;i c&#7913; h&#7887;i, ch&#432;a ch&#7855;c c&oacute; ng&#432;&#7901;i &#273;&aacute;p...") & " tabindex=""-1""><acronym title=""H&#7887;i c&#7913; h&#7887;i, ch&#432;a ch&#7855;c c&oacute; ng&#432;&#7901;i &#273;&aacute;p..."">H&#7887;i / &#273;&aacute;p</acronym></a>" & _
			"</font>"
end sub

if strGroupCategories = "1" then
	if Session(strCookieURL & "GROUP_NAME") = "" then
		GROUPNAME = " Default Groups "
	else
		GROUPNAME = Session(strCookieURL & "GROUP_NAME")
	end if
	'Forum_SQL - Get Groups
	strSql = "SELECT GROUP_ID, GROUP_CATID " 
	strSql = strSql & " FROM " & strTablePrefix & "GROUPS "
	strSql = strSql & " WHERE GROUP_ID = " & Group
	set rsgroups = Server.CreateObject("ADODB.Recordset")
	rsgroups.Open strSql, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
	if rsgroups.EOF then
		recGroupCatCount = ""
	else
		allGroupCatData = rsgroups.GetRows(adGetRowsRest)
		recGroupCatCount = UBound(allGroupCatData, 2)
	end if
	rsgroups.Close
	set rsgroups = nothing
end if
%>

<style>
	.TieuDe {color: #FFFF00;
	font-weight: bold;
	font-size: 30px;
	}

	.MenuBar
	{
		background-color:#990000;
		width:100%;
	}
    .Navigation
    {
	color:White;
	font-family:Arial;
	font-weight:bold;
	font-size:11px;
	background-color:#FFCC00;
    }
    .Navigation a
    {
	color:#FF0000;
	font-family:Arial;
	font-weight:bold;
	font-size:14px;
	text-decoration: none;
   	}
   	.Navigation a:hover
    {
	color:#FFFFFF;
    }
	
	body
	{
		color:#333333;
    	font-family:Arial; 
    	font-size:12px;
	}
	.Title
	{
		background-image:url(Banner.jpg) ;
		background-color:#FF9900;
		overflow:visible;
	    font-size: 20px;
		text-transform:uppercase;
		color:#0066FF;
		vertical-align:middle;
		text-align:center;
		padding-top:20px;
		height:80px;
	}
	.TableList
	{
		width:100%;
	}
	
	.TableList th
	{
		background-color:#666666;
		color:white;
		border-color:#666666;
		font-weight:bold; 
    	font-size:12px;
		padding:5px 0px 5px 0px;
	}
	.TableList td
	{
		border-bottom-color:#333333;
		border-bottom-style:dotted;
		border-bottom-width:thin;
		padding:0px 0px 0px 0px;
	}
</style>
</td></tr></table>
<script type="text/javascript" src="stmenu.js"></script>
<span>
<script id="sothink_widgets:dwwidget_dhtmlmenu4_29_2008_1.pgt" type="text/javascript">
<!--
stm_bm(["menu61b2",850,"","blank.gif",0,"","",0,0,250,0,1000,1,0,0,"","540",0,0,1,2,"default","hand",""],this);
stm_bp("p0",[0,4,0,0,0,0,18,28,100,"",-2,"",-2,50,0,0,"#799BD8","transparent","060420line.gif",3,0,0,"#000000","",-1,-1,0,"transparent","",0,"060420lineb.gif",-1,-1,0,"transparent","",0,"",-1,-1,0,"transparent","",0,"060420linea.gif",-1,-1,0,"transparent","",0,"","","","",0,0,0,0,0,0,0,0]);
stm_aix("p0i0","p0i0",[0,"Danh sách đội bóng  ","","",-1,-1,0,"teams.asp","_self","","","060420icon.gif","060420icon1.gif",20,25,0,"","",0,0,0,0,1,"#FFFFF7",1,"#B5BED6",1,"","",2,3,0,0,"#FFFFF7","#000000","#FFFFFF","#FFFF00","bold 9pt Verdana","bold 9pt Verdana",0,0],90,10);
stm_aix("p0i1","p0i0",[0,"Lịch thi đấu  ","","",-1,-1,0,"fixtures.asp","_self","","","060420icon.gif","060420icon1.gif",20,25,0,"","",0,0,0,0,1,"#FFFFF7",1,"#B5BED6",1,"","",3],90,23);
stm_aix("p0i2","p0i1",[0,"Chọn đội","","",-1,-1,0,"chondoi.asp"],90,23);
stm_aix("p0i3","p0i1",[0,"Thông tin chọn đội  ","","",-1,-1,0,"Results.asp"],90,23);
stm_aix("p0i4","p0i1",[0,"Bảng phong thần  ","","",-1,-1,0,"ResultsAll.asp"],90,23);
stm_aix("p0i5","p0i1",[0,"Thành tích cá nhân  ","","",-1,-1,0,"ResultsAllRev.asp"],90,23);
//stm_aix("p0i6","p0i1",[0,"Liên kết","","",-1,-1,0,"#","_self","","","060420icon.gif","060420icon1.gif",20,25,0,"060420arrow.gif","060420arrow1.gif",28,5],110,23);
//stm_bp("p1",[1,4,0,0,3,0,18,0,100,"",-2,"",-2,80,0,0,"#799BD8","#0B3949","",3,1,1,"#000000"]);
//stm_aix("p1i0", "p0i0", [0, " Trang chủ Worlcup 2010", "", "", -1, -1, 0, "http://www.fifa.com/worldcup/", "_blank", "", "", "060420icon3.gif", "060420icon2.gif", 18, 16, 0, "", "", 0, 0, 0, 0, 1, "#FFFFF7", 1, "#680000", 0, "", "", 3, 0, 1, 1, "#FFFFCC", "#CCCC00", "#00FFFF", "#FFFF00", "bold 8pt Verdana", "bold 8pt Verdana"], 200, 18);
//stm_aix("p1i1","p1i0",[0," Live Score","","",-1,-1,0,"http://livescore.com/"],200,18);
//stm_aix("p1i2","p1i0",[0," Asianbookie","","",-1,-1,0,"http://www.asianbookie.com/index.cfm?league=4&tz=7"],200,18);
//stm_aix("p1i3", "p1i0", [0, " Yahoo Sport", "", "", -1, -1, 0, "http://sports.yahoo.com/soccer/world-cup/"], 200, 18);
//stm_aix("p1i4","p1i0",[0," VNEpress","","",-1,-1,0,"http://vnexpress.net/Vietnam/The-thao/"],200,18);
stm_ep();
stm_ep();
stm_em();
//-->
</script>
</span>