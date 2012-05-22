<%
'    Bulletin Board Service
'    Copyright (C) 2000 Sudaraka Wijesinghe <sudaraka.wijesinghe@gmail.com>
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with this program.  If not, see <http://www.gnu.org/licenses/>.

response.expires=0
if session("user")="" then
	response.write("<font face=""verdana"" size=1><b>ERROR:</b> Session profile is not available.<br>Please contact <a href=""mailto:sys@post.com?subject=BBS Error""> Systems</a>.<hr>")
	response.end
end if
%>
<html>
<head>
<title> BBS - Compose</title>
<link href="/<%=session("homepath")%>main.css" rel=stylesheet>
</head>
<body vlink=#0000ff><font face="arial,verdana" size=2>
<font face="times new roman" size=7><i><%out3d "Bulletin Board Service.",0,0%></i></font>
<form action="/<%=session("homepath")%>docomp.asp" method="POST">
<table style="position:absolute;top:75;" width=100% bgcolor=#6699cc bordercolordark=#000000 bordercolorlight=#99ccff class=blue_table>
<tr>
<td align=right><font face="arial,verdana" size=2><B>Author:</td>
<td width=10></td>
<td><input type="text" size="63" name="author" class=blue_textbox value="<%=session("fullname")%>"></td>
</tr>
<%if (session("acclvl")and 32)=32 then%>
<tr>
<td align=right><font face="arial,verdana" size=2><B>To:</td>
<td width=10></td>
<td><input type="text" size="63" name="owner" class=blue_textbox></td>
</tr>
<%else%>
<input type="hidden" name="owner" value="public_bbs">
<%end if%>
<tr>
<td align=right><font face="arial,verdana" size=2><B>Title:</td>
<td width=10></td>
<td><input type="text" size="63" name="title" class=blue_textbox></td>
</tr>
<tr>
<td align=right valign=top><font face="arial,verdana" size=2><B>Message:</td>
<td width=10></td>
<td><textarea name="message" rows="6" cols="62" class=blue_msgbox></textarea></td>
</tr>
<%if (session("acclvl")and 32)=32 and session("sign")<>"" then%>
<tr>
<td align=right valign=top><font face="arial,verdana" size=2><B><input type="checkbox" name="sign" value="1" checked>Sign message:</td>
<td width=10></td>
<td><font face="comic sans ms" size=2 color=#ffffff>
<%
fmtmsg=""
tmpstr=session("sign")
while instr(tmpstr,chr(10))>0
	fmtmsg=fmtmsg&splitstr(tmpstr,chr(10))&"<br>"
wend
fmtmsg=fmtmsg&tmpstr
tmpstr=fmtmsg
fmtmsg=""
while instr(tmpstr,"  ")>0
	fmtmsg=fmtmsg&splitstr(tmpstr,"  ")&"&nbsp;&nbsp;"
wend
response.write(fmtmsg&tmpstr)
%></td>
</tr>
<%else%>
<input type="hidden" name="sign" value="0">
<%end if%>
<tr>
<td></td><td></td><td></td>
</tr>
<tr>
<td></td>
<td width=10></td>
<td><input type="submit" value="Post" class=blue_button> <input type="reset" value="Reset form" class=blue_button> <input type="button" value="Go Back" class=blue_button onclick="location.href='/<%=session("homepath")%>';"></td>
</tr>
</table>
</form>
<!--#include virtual="/bbs/codebase.asp"-->
