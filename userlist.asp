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
<title> BBS - User list</title>
<link href="/<%=session("homepath")%>main.css" rel=stylesheet>
</head>
<body vlink=#0000ff><font face="arial,verdana" size=2>
<font face="times new roman" size=7><i><%out3d "Bulletin Board Service.",0,0%></i></font>
&nbsp;<p><br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Following list contains the general information of current users of the  BBS. You can click on the full name to contact them via e-mail or compose a BBS message to the user name.<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Users currently online are marked with a "<font color=#ff0000>*</font>".<p>
<div style="position:absolute;left:3;">
<table class=blue_table width=100% bgcolor=#c0c0c0 bordercolordark=#808080 bordercolorlight=#ffffff cellspacing=0><tr>
<td align=center width=5%></td>
<td class=listhead width=25%>Username</td>
<td class=listhead>Full name</td>
</tr>
</table>
<tr>
<table width=100% cellspacing=0>
<%
set con=server.createobject("adodb.connection")
con.open "driver={microsoft access driver (*.mdb)};dbq="&server.mappath("/"&session("dbpath")&"bulletin.mdb")
set rs=con.execute("select * from profiles")
rc=1
while not rs.eof
%>
<tr<%if(rc mod 2)=0 then response.write(" bgcolor=#eeeeee")%>>
<td align=center width=5%><font face="arial" size=2><b><%=rc%>.</td>
<td width=25%><font face="arial" size=2><%response.write(rs.fields("username").value)%></td>
<td><font face="verdana,arial" size=1>
<%
if rs.fields("email").value<>"" then
	response.write("<a style=""text-decoration:none;"" href=""mailto:"&rs.fields("email").value&""">")
end if
response.write(rs.fields("fullname").value)
if rs.fields("email").value<>"" then
	response.write("</a>")
end if
rs.movenext
rc=rc+1
wend
rs.close
con.close
set rs=nothing
set con=nothing
%>
</table><br>
<hr>
<a href="/<%=session("homepath")%>"><b>Go Back</b></a>
</div>
</body>
</html>
<!--#include virtual="/bbs/codebase.asp"-->
