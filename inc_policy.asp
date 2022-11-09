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

Response.Write	"      <table width=""100%"" border=""0"">" & vbNewLine & _
		"        <tr>" & vbNewLine & _
		"          <td><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
		"          " & getCurrentIcon(strIconFolderOpen,"","") & "&nbsp;<a href=""default.asp"">All Forums</a><br />" & vbNewLine & _
		"          " & getCurrentIcon(strIconBar,"","") & getCurrentIcon(strIconFolderOpenTopic,"","") & "&nbsp;Quy Lu&#7853;t Di&#7877;n &ETH;&agrave;n</font></td>" & vbNewLine & _
		"        </tr>" & vbNewLine & _
		"      </table>" & vbNewLine

if strProhibitNewMembers <> "1" then
	Response.Write	"      <table border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""0"" align=""center"">" & vbNewLine & _
		"        <tr>" & vbNewLine & _
		"          <td bgcolor=""" & strTableBorderColor & """>" & vbNewLine & _
		"            <table border=""0"" width=""100%"" cellspacing=""1"" cellpadding=""4"">" & vbNewLine & _
		"              <tr>" & vbNewLine & _
		"                <td bgcolor=""" & strCategoryCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """><b>&nbsp;Quy Lu&#7853;t Di&#7877;n &ETH;&agrave;n&nbsp; " & strForumTitle & "</b></font></td>" & vbNewLine & _
		"              </tr>" & vbNewLine & _
		"              <tr>" & vbNewLine & _
		"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
		"                <p>Vi&#7879;c &#273;&#259;ng k&iacute; v&agrave;o di&#7877;n d&agrave;n ho&agrave;n to&agrave;n mi&#7877;n ph&iacute;.</p>" & vbNewLine & _
		"                <p>Tr&#432;&#7899;c khi &#273;&#259;ng k&iacute; t&#7845;t c&#7843; &#273;&#7873;u ph&#7843;i &#273;&#7891;ng &yacute; v&#7899;i nh&#7919;ng quy &#432;&#7899;c sau :</p>" & vbNewLine & _
		"				 <blockquote>"& _
		"                <p>* T&#7845;t c&#7843; b&agrave;i g&#7903;i &#273;&#7873;u ph&#7843;i vi&#7871;t b&#7857;ng ti&#7871;ng Vi&#7879;t &#273;&#7847;y &#273;&#7911;  d&#7845;u n&#7871;u kh&ocirc;ng s&#7869; b&#7883; x&oacute;a b&#7887;.<br>"& _
		"				 * Xin &#273;&#7915;ng b&agrave;n v&#7873; chuy&#7879;n ch&iacute;nh tr&#7883;, t&ocirc;n gi&aacute;o hay nh&#7919;ng &#273;&#7873; t&agrave;i d&#7877; g&acirc;y hi&#7873;m kh&iacute;ch.<br>"& _
		"				 * Kh&ocirc;ng &#273;&#432;&#7907;c k&#7923; th&#7883; ch&#7911;ng t&#7897;c, gi&#7899;i t&iacute;nh, n&#417;i c&#432; ng&#7909; hay tu&#7893;i t&aacute;c.<br>"& _
		"				 * Xin h&atilde;y c&#432; x&#7917; h&ograve;a nh&atilde; v&#7899;i nhau! Ch&uacute;ng t&ocirc;i s&#7869; kh&ocirc;ng ch&#7845;p nh&#7853;n b&#7845;t c&#7913; h&igrave;nh th&#7913;c c&ocirc;ng k&iacute;ch, m&#7841; l&#7883; n&agrave;o nh&#7855;m v&agrave;o m&#7897;t c&aacute; nh&acirc;n hay m&#7897;t &#273;o&agrave;n th&#7875;.<br>"& _
		"				 * Xin &#273;&#7915;ng g&#7903;i b&agrave;i ho&#7863;c h&igrave;nh &#7843;nh c&oacute; n&#7897;i dung th&ocirc; t&#7909;c, khi&ecirc;u d&acirc;m.</p>" & vbNewLine & _
		"				 </blockquote>"& vbNewLine & _
		"                <p>Ti&ecirc;u ch&iacute; khi tham gia &#273;&#259;ng k&yacute; :</p>" & vbNewLine & _
		"				 <blockquote>"& _
		"                <p>* Khi &#273;&#259;ng k&yacute; ph&#7843;i &#273;i&#7873;n &#273;&#7847;y &#273;&#7911; c&aacute;c th&ocirc;ng tin li&ecirc;n quan &#273;&#7871;n ng&#432;&#7901;i ch&#417;i.<br>"& _
		"				 * H&#7885; t&ecirc;n, s&#7889; &#273;i&#7879;n tho&#7841;i v&agrave; Email ph&#7843;i l&agrave; nh&#7919;ng th&ocirc;ng tin ch&iacute;nh x&aacute;c &#273;&#7875; d&#7877; b&#7873; &#273;&ograve;i n&#7907; (tr&aacute;nh t&igrave;nh tr&#7841;ng n&#7907; kh&oacute; &#273;&ograve;i).<br>"& _
		"				 * T&#7845;t c&#7843; v&igrave; m&#7897;t m&#7909;c ti&ecirc;u chung: <em><strong>&quot;Vui L&agrave; Ch&iacute;nh - Nh&#7853;u l&agrave; M&#432;&#7901;i &quot;</strong></em>.</p>" & vbNewLine & _
		"				 </blockquote>"& vbNewLine & _
		"                <p>N&#7871;u vi ph&#7841;m ch&uacute;ng t&ocirc;i s&#7869; t&#432;&#7899;c quy&#7873;n l&#7907;i th&agrave;nh vi&ecirc;n m&agrave; kh&ocirc;ng c&#7847;n ph&#7843;i b&aacute;o tr&#432;&#7899;c. &ETH;&acirc;y l&agrave; m&#7897;t di&#7877;n &#273;&agrave;n c&oacute; t&iacute;nh c&aacute;ch c&aacute; nh&acirc;n v&igrave; v&#7853;y n&#7897;i dung b&agrave;i g&#7903;i ch&uacute;ng t&ocirc;i s&#7869; kh&ocirc;ng ch&#7883;u m&#7897;t tr&aacute;ch nhi&#7879;m n&agrave;o.</p>" & vbNewLine & _
		"                <hr size=""1"">" & vbNewLine & _
		"                  <table align=""center"" border=""0"">" & vbNewLine & _
		"                    <tbody>" & vbNewLine & _
		"                      <tr>" & vbNewLine & _
		"                        <td>" & vbNewLine & _
		"                        <form action=""register.asp?mode=Register"" id=""form1"" method=""post"" name=""form1"">" & vbNewLine & _
		"                        <input name=""Refer"" type=""hidden"" value=""" & strReferer & """>" & vbNewLine & _
		"                        <input name=""policy_accept"" type=""hidden"" value=""true"">" & vbNewLine & _
		"                        <input name=""Submit"" type=""Submit"" value=""&#272;&#7891;ng &yacute;"">" & vbNewLine & _
		"                        </form>" & vbNewLine & _
		"                        </td>" & vbNewLine & _
		"                        <td>" & vbNewLine & _
		"                        <form action=""JavaScript:history.go(-1)"" id=""form2"" method=""post"" name=""form2"">" & vbNewLine & _
		"                        <input name=""Submit"" type=""Submit"" value=""Kh&ocirc;ng &#273;&#7891;ng &yacute;"">" & _
		"                        </form>" & _
		"                        </td>" & _
		"                      </tr>" & _
		"                    </tbody>" & _
		"                  </table>" & _
		"                <hr size=""1"">" & vbNewLine & _
		"                <p>M&#7885;i th&#7855;c m&#7855;c v&agrave; khi&#7871;u n&#7841;i vui l&ograve;ng li&ecirc;n h&#7879; :<br>- Phone: 0913.777.426<br>- Email: " & _
		"				 <span class=""spnMessageText""><a href=""mailto:" & strSender & """>" & strSender & "</a></span></p>" & vbNewLine & _
		"                </font></td>" & vbNewLine & _
		"              </tr>" & vbNewLine & _
		"            </table>" & vbNewLine & _
		"          </td>" & vbNewLine & _
		"        </tr>" & vbNewLine & _
		"      </table>" & vbNewLine & _
		"      <br />" & vbNewLine
else
	Response.Write	"    <br /><p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """>Sorry, we are not accepting any new Members at this time.</font></p>" & vbNewLine & _
		"    <meta http-equiv=""Refresh"" content=""5; URL=default.asp"">" & vbNewLine & _
		"    <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""default.asp"">Back To Forum</font></a></p><br />" & vbNewLine
end if
WriteFooter
Response.End
%>
