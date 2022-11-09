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
<!--#INCLUDE FILE="config.asp"-->
<!--#INCLUDE FILE="inc_sha256.asp"-->
<!--#INCLUDE FILE="inc_header.asp" -->
<!--#INCLUDE FILE="inc_func_admin.asp" -->
<!--#INCLUDE FILE="inc_func_member.asp" -->
<%
if Session(strCookieURL & "Approval") <> "15916941253" Then 
	scriptname = split(request.servervariables("SCRIPT_NAME"),"/")
	Response.Redirect "admin_login.asp?target=" & scriptname(ubound(scriptname))
end if
Response.Write	"      <table border=""0"" width=""100%"">" & vbNewLine & _
		"        <tr>" & vbNewLine & _
		"          <td width=""33%"" align=""left"" nowrap><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
		"          " & getCurrentIcon(strIconFolderOpen,"","") & "&nbsp;<a href=""default.asp"">All&nbsp;Forums</a><br />" & vbNewLine & _
		"          " & getCurrentIcon(strIconBar,"","") & getCurrentIcon(strIconFolderOpen,"","") & "&nbsp;<a href=""admin_home.asp"">Admin&nbsp;Section</a><br />" & vbNewLine & _
		"          " & getCurrentIcon(strIconBlank,"","") & getCurrentIcon(strIconBar,"","") & getCurrentIcon(strIconFolderOpenTopic,"","") & "&nbsp;Member&nbsp;Details&nbsp;Configuration<br /></font></td>" & vbNewLine & _
		"        </tr>" & vbNewLine & _
		"      </table>" & vbNewLine
if Request.Form("Method_Type") = "Write_Configuration" then 
	Err_Msg = ""
	
	if Request.Form("strAge") = "1" and Request.Form("strAgeDOB") = "1" then 
		Err_Msg = Err_Msg & "<li>Age and Birth Date cannot both be On at the same time</li>"
	end if
	intAge = ChkString(trim(Request.Form("strMinAge")), "SQLString")
	if len(intAge) = 0 then
		intAge = 0
	end if
	if not isNumeric(intAge) then
		Err_Msg = Err_Msg & "<li>Minimum Age must be a numerical value.</li>"
	end if

	if Err_Msg = "" then
		for each key in Request.Form 
			if left(key,3) = "str" then 
				strDummy = SetConfigValue(1, key, ChkString(Request.Form(key),"SQLString"))
			end if
		next

		Application(strCookieURL & "ConfigLoaded") = ""

		Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>Configuration Posted!</font></p>" & vbNewLine & _
				"      <meta http-equiv=""Refresh"" content=""2; URL=admin_home.asp"">" & vbNewLine & _
				"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>Congratulations!</font></p>" & vbNewLine & _
				"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""admin_home.asp"">Back To Admin Home</font></a></p>" & vbNewLine
	else
		Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """>There Was A Problem With Your Details</font></p>" & vbNewLine & _
				"      <table align=""center"" border=""0"">" & vbNewLine & _
				"        <tr>" & vbNewLine & _
				"          <td><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHiLiteFontColor & """><ul>" & Err_Msg & "</ul></font></td>" & vbNewLine & _
				"        </tr>" & vbNewLine & _
				"      </table>" & vbNewLine & _
				"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""JavaScript:history.go(-1)"">Go Back To Enter Data</a></font></p>" & vbNewLine
	end if
else
	Response.Write	"      <form action=""admin_config_members.asp"" method=""post"" id=""Form1"" name=""Form1"">" & vbNewLine & _
			"      <input type=""hidden"" name=""Method_Type"" value=""Write_Configuration"">" & vbNewLine & _
			"      <table border=""0"" cellspacing=""0"" cellpadding=""0"" align=""center"">" & vbNewLine & _
			"        <tr>" & vbNewLine & _
			"          <td bgcolor=""" & strPopUpBorderColor & """>" & vbNewLine & _
			"            <table border=""0"" cellspacing=""1"" cellpadding=""1"" width=""100%"">" & vbNewLine & _
			"              <tr valign=""top"">" & vbNewLine & _
			"                <td bgcolor=""" & strHeadCellColor & """ colspan=""2""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """><b>Member Details Configuration</b></font></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""top"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Fullname:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                On: <input type=""radio"" class=""radio"" name=""strFullName"" value=""1""" & chkRadio(strFullName,0,false) & ">&nbsp;" & vbNewLine & _
			"                Off: <input type=""radio"" class=""radio"" name=""strFullName"" value=""0""" & chkRadio(strFullName,0,true) & ">" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=members#FullName')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""top"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Picture:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                On: <input type=""radio"" class=""radio"" name=""strPicture"" value=""1""" & chkRadio(strPicture,0,false) & ">&nbsp;" & vbNewLine & _
			"                Off: <input type=""radio"" class=""radio"" name=""strPicture"" value=""0""" & chkRadio(strPicture,0,true) & ">" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=members#Picture')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""top"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Recent Topics:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                On: <input type=""radio"" class=""radio"" name=""strRecentTopics"" value=""1""" & chkRadio(strRecentTopics,0,false) & ">&nbsp;" & vbNewLine & _
			"                Off: <input type=""radio"" class=""radio"" name=""strRecentTopics"" value=""0""" & chkRadio(strRecentTopics,0,true) & ">" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=members#RecentTopics')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""top"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Gender (male/female):</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                On: <input type=""radio"" class=""radio"" name=""strSex"" value=""1""" & chkRadio(strSex,0,false) & ">&nbsp;" & vbNewLine & _
			"                Off: <input type=""radio"" class=""radio"" name=""strSex"" value=""0""" & chkRadio(strSex,0,true) & ">" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=members#Sex')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""top"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Age:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                On: <input type=""radio"" class=""radio"" name=""strAge"" value=""1""" & chkRadio(strAge,0,false) & ">&nbsp;" & vbNewLine & _
			"                Off: <input type=""radio"" class=""radio"" name=""strAge"" value=""0""" & chkRadio(strAge,0,true) & ">" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=members#Age')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""top"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Birth Date:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                On: <input type=""radio"" class=""radio"" name=""strAgeDOB"" value=""1""" & chkRadio(strAgeDOB,0,false) & ">&nbsp;" & vbNewLine & _
			"                Off: <input type=""radio"" class=""radio"" name=""strAgeDOB"" value=""0""" & chkRadio(strAgeDOB,0,true) & ">" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=members#AgeDOB')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr>" & vbNewLine & _
			"                <td bgcolor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Minimum Age:</b>&nbsp;</font>" & vbNewLine
	intYoungest = 0
	if strAge = "1" then
		set rs = my_Conn.execute(TopSQL("SELECT M_AGE FROM " & strMemberTablePrefix & "MEMBERS WHERE M_AGE <> '' AND M_STATUS <> 0 ORDER BY M_AGE ASC", 1))
		if not rs.eof then
			intYoungest = cInt(rs("M_AGE"))
		end if
		rs.close
		set rs = nothing
	elseif strAgeDOB = "1" then
		set rs = my_Conn.execute(TopSQL("SELECT M_DOB FROM " & strMemberTablePrefix & "MEMBERS WHERE M_DOB <> '' AND M_STATUS <> 0 ORDER BY M_DOB DESC", 1))
		if not rs.eof then
			intYoungest = cInt(DisplayUsersAge(DOBToDate(rs("M_DOB"))))
		end if
		rs.close
		set rs = nothing
	end if
	if intYoungest > 0 then
		Response.Write	"                <font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strHiLiteFontColor & """><br><b>Youngest member is " & intYoungest & "&nbsp;</b></font>" & vbNewLine
	end if
	Response.Write	"                </td>" & vbNewLine & _
			"                <td bgcolor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><input type=""text"" name=""strMinAge"" value="""&strMinAge&""" maxlength=""2"" size=""16"" /> " & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=members#MinAge')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></font></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""top"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>City:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                On: <input type=""radio"" class=""radio"" name=""strCity"" value=""1""" & chkRadio(strCity,0,false) & ">&nbsp;" & vbNewLine & _
			"                Off: <input type=""radio"" class=""radio"" name=""strCity"" value=""0""" & chkRadio(strCity,0,true) & ">" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=members#City')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""top"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>State:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                On: <input type=""radio"" class=""radio"" name=""strState"" value=""1""" & chkRadio(strState,0,false) & ">&nbsp;" & vbNewLine & _
			"                Off: <input type=""radio"" class=""radio"" name=""strState"" value=""0""" & chkRadio(strState,0,true) & ">" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=members#State')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""top"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Country:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                On: <input type=""radio"" class=""radio"" name=""strCountry"" value=""1""" & chkRadio(strCountry,0,false) & ">&nbsp;" & vbNewLine & _
			"                Off: <input type=""radio"" class=""radio"" name=""strCountry"" value=""0""" & chkRadio(strCountry,0,true) & ">" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=members#Country')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""top"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>AIM:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                On: <input type=""radio"" class=""radio"" name=""strAIM"" value=""1""" & chkRadio(strAIM,0,false) & ">&nbsp;" & vbNewLine & _
			"                Off: <input type=""radio"" class=""radio"" name=""strAIM"" value=""0""" & chkRadio(strAIM,0,true) & ">" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=members#aim')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""top"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>ICQ:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                On: <input type=""radio"" class=""radio"" name=""strICQ"" value=""1""" & chkRadio(strICQ,0,false) & ">&nbsp;" & vbNewLine & _
			"                Off: <input type=""radio"" class=""radio"" name=""strICQ"" value=""0""" & chkRadio(strICQ,0,true) & ">" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=members#icq')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""top"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>MSN:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                On: <input type=""radio"" class=""radio"" name=""strMSN"" value=""1""" & chkRadio(strMSN,0,false) & ">&nbsp;" & vbNewLine & _
			"                Off: <input type=""radio"" class=""radio"" name=""strMSN"" value=""0""" & chkRadio(strMSN,0,true) & ">" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=members#msn')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""top"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>YAHOO:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                On: <input type=""radio"" class=""radio"" name=""strYAHOO"" value=""1""" & chkRadio(strYAHOO,0,false) & ">&nbsp;" & vbNewLine & _
			"                Off: <input type=""radio"" class=""radio"" name=""strYAHOO"" value=""0""" & chkRadio(strYAHOO,0,true) & ">" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=members#yahoo')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""top"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Occupation:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                On: <input type=""radio"" class=""radio"" name=""strOccupation"" value=""1""" & chkRadio(strOccupation,0,false) & ">&nbsp;" & vbNewLine & _
			"                Off: <input type=""radio"" class=""radio"" name=""strOccupation"" value=""0""" & chkRadio(strOccupation,0,true) & ">" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=members#Occupation')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""top"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Homepages:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                On: <input type=""radio"" class=""radio"" name=""strHomepage"" value=""1""" & chkRadio(strHomepage,0,false) & ">&nbsp;" & vbNewLine & _
			"                Off: <input type=""radio"" class=""radio"" name=""strHomepage"" value=""0""" & chkRadio(strHomepage,0,true) & ">" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=members#Homepages')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""top"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Favorite Links:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                On: <input type=""radio"" class=""radio"" name=""strFavLinks"" value=""1""" & chkRadio(strFavLinks,0,false) & ">&nbsp;" & vbNewLine & _
			"                Off: <input type=""radio"" class=""radio"" name=""strFavLinks"" value=""0""" & chkRadio(strFavLinks,0,true) & ">" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=members#FavLinks')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""top"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Marital Status:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                On: <input type=""radio"" class=""radio"" name=""strMarStatus"" value=""1""" & chkRadio(strMarStatus,0,false) & ">&nbsp;" & vbNewLine & _
			"                Off: <input type=""radio"" class=""radio"" name=""strMarStatus"" value=""0""" & chkRadio(strMarStatus,0,true) & ">" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=members#MStatus')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""top"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Bio:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                On: <input type=""radio"" class=""radio"" name=""strBio"" value=""1""" & chkRadio(strBio,0,false) & ">&nbsp;" & vbNewLine & _
			"                Off: <input type=""radio"" class=""radio"" name=""strBio"" value=""0""" & chkRadio(strBio,0,true) & ">" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=members#Bio')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""top"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Hobbies:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                On: <input type=""radio"" class=""radio"" name=""strHobbies"" value=""1""" & chkRadio(strHobbies,0,false) & ">&nbsp;" & vbNewLine & _
			"                Off: <input type=""radio"" class=""radio"" name=""strHobbies"" value=""0""" & chkRadio(strHobbies,0,true) & ">" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=members#hobbies')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""top"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Latest News:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                On: <input type=""radio"" class=""radio"" name=""strLNews"" value=""1""" & chkRadio(strLNews,0,false) & ">&nbsp;" & vbNewLine & _
			"                Off: <input type=""radio"" class=""radio"" name=""strLNews"" value=""0""" & chkRadio(strLNews,0,true) & ">" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=members#LNews')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""top"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Quote:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                On: <input type=""radio"" class=""radio"" name=""strQuote"" value=""1""" & chkRadio(strQuote,0,false) & ">&nbsp;" & vbNewLine & _
			"                Off: <input type=""radio"" class=""radio"" name=""strQuote"" value=""0""" & chkRadio(strQuote,0,true) & ">" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=members#Quote')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""top"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ colspan=""2"" align=""center""><input type=""submit"" value=""Submit New Config"" id=""submit1"" name=""submit1""> <input type=""reset"" value=""Reset Old Values"" id=""reset1"" name=""reset1""></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"            </table>" & vbNewLine & _
			"          </td>" & vbNewLine & _
			"        </tr>" & vbNewLine & _
			"      </table>" & vbNewLine & _
			"      </form>" & vbNewLine
end if 
WriteFooter
%>
