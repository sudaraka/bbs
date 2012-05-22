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
<title> BBS</title>
<link href="/<%=session("homepath")%>main.css" rel=stylesheet>
</head>
<body vlink=#0000ff><font face="arial,verdana" size=2>
<font face="times new roman" size=7><i><%out3d "Bulletin Board Service.",0,0%></i></font>
<div style="position:absolute;top:55;">
<%
set con=server.createobject("adodb.connection")
con.open "driver={microsoft access driver (*.mdb)};dbq="&server.mappath("/"&session("dbpath")&"bulletin.mdb")
set rs=con.execute("select * from bbsdata where id="&request.querystring("msgid"))
if not rs.eof then
	if lcase(rs.fields("owner"))=lcase(session("user")) then
%>
<table width=100%>
<tr>
<td align=right width=10%><font face="arial,verdana" size=2><b>From:</td>
<td><font face="arial,verdana" size=2><%=rs.fields("author").value%></td>
</tr>
<tr>
<td align=right width=10%><font face="arial,verdana" size=2><b>Title:</td>
<td><font face="arial,verdana" size=2><b><%=rs.fields("title").value%></td>
</tr>
<tr>
<td align=right width=10%><font face="arial,verdana" size=2><b>Date:</td>
<td><font face="arial,verdana" size=2><%=rs.fields("date").value%></td>
</tr>
</table>
<table width=100%>
<th bgcolor=#808080 class=blue_table><font face="arial,verdana" size=2 color=#ffffff>- Message -</th>
<%
fmtmsg=""
tmpstr=rs.fields("message").value
while instr(tmpstr,chr(10))>0
	fmtmsg=fmtmsg&splitstr(tmpstr,chr(10))&"<br>"
wend
fmtmsg=fmtmsg&tmpstr
tmpstr=fmtmsg
fmtmsg=""
while instr(tmpstr,"  ")>0
	fmtmsg=fmtmsg&splitstr(tmpstr,"  ")&"&nbsp;&nbsp;"
wend
fmtmsg=fmtmsg&tmpstr
%>
<tr>
<td><font face="comic sans ms,verdana,courier new,arial" size=2><%=fmtmsg%></td>
</tr>
</table>
<p>
<%
	else
%>
<font face="arial,verdana" size=2 color=#ff0000><b>Access denied.</font>
<%
	end if
else
%>
<font face="arial,verdana" size=2><b>Message not found.</font><br>
<%end if%>
<table width=100% bgcolor=#6699cc bordercolordark=#000000 bordercolorlight=#99ccff class=blue_table>
<tr>
<td align=center>
<%if lcase(session("user"))<>"public_bbs" and session("user")<>"" and not rs.eof then%>
<form name="delform" method="POST" action="/<%=session("homepath")%>delmsg.asp">
<input type="hidden" value="<%=request.querystring("msgid")%>" name="dmsg">
<input type="submit" value="Delete" class=blue_button>
<%end if%>
<input type="button" value="Go Back" class=blue_button onclick="location.href='/<%=session("homepath")%>default.asp';"></td>
<%if lcase(session("user"))<>"public_bbs" and session("user")<>"" and not rs.eof then%>
</form>
<%
end if
rs.close
con.close
set rs=nothing
set con=nothing
%>
</tr>
</table>
</div>
</body>
</html>
<!--#include virtual="/bbs/codebase.asp"-->
