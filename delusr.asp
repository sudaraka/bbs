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

if session("user")="" or (session("acclvl")and 16)<>16 or session("user")="public_bbs" then
	response.write("<font face=""verdana"" size=1><b>ERROR:</b> Session profile is not available.<br>Please contact <a href=""mailto:sys@post.com?subject=BBS Error""> Systems</a>.<hr>")
	response.end
end if
%>
<html>
<head>
<title> BBS - User list</title>
<link href="/<%=session("homepath")%>main.css" rel=stylesheet>
</head>
<body vlink=#0000ff><font face="arial,verdana" size=2>
<font face="times new roman" size=7><i><%out3d "Bulletin Board Service.",0,0%></i></font>
&nbsp;<p><br>
<center>
<form method=post action="/<%=session("homepath")%>dodelusr.asp">
<input type="hidden" name="delusr" value="<%=request.querystring("userid")%>">
<table bgcolor=#6699cc bordercolordark=#000000 bordercolorlight=#99ccff class=blue_table>
<tr>
<td align=center><font face="arial" size=2>Are you sure you want to delete <%=request.querystring("username")%>'s account?<br><font face="comic sans ms" color=#c00000 size=1>(WARNING!, Once deleted account is lost forever.)</font><p>
<input type="submit" class=blue_button value="Yes">
<input type="button" class=blue_button value="No " onclick="location.href='/<%=session("homepath")%>config.asp';">
</td>
</tr></table></form>
</body>
</html>
<!--#include virtual="/bbs/codebase.asp"-->
