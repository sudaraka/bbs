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
if session("user")="" or session("acclvl")=0 or session("user")="public_bbs" then
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
<div style="position:absolute;top:55;left:3;">
<table width=100% bgcolor=#6699cc bordercolordark=#000000 bordercolorlight=#99ccff class=blue_table>
<tr>
<td align=center>
<form name="delform" method="POST" action="/<%=session("homepath")%>delpub.asp">
<input type="submit" value="Delete" class=blue_button>
<input type="button" value="Go Back" class=blue_button onclick="location.href='/<%=session("homepath")%>config.asp';"></td>
</tr>
</table>
<table class=blue_table width=100% bgcolor=#c0c0c0 bordercolordark=#808080 bordercolorlight=#ffffff cellspacing=0><tr>
<td width=2%><input type="checkbox" name="chal" onclick="selall();"></td>
<td class=listhead align=center>Title</td>
<td class=listhead width=25%>Date</td>
<td class=listhead width=25% align=right>Author</td>
</tr>
</table>
<tr>
<table width=100% cellspacing=0>
<%
set con=server.createobject("adodb.connection")
con.open "driver={microsoft access driver (*.mdb)};dbq="&server.mappath("/"&session("dbpath")&"bulletin.mdb")
set rs=con.execute("select * from bbsdata where owner='public_bbs'")
rc=1
while not rs.eof
%>
<tr<%if(rc mod 2)=0 then response.write(" bgcolor=#eeeeee")%>>
<td align=center width=2%><input type="checkbox" name="dmsg" value="<%=rs.fields("id").value%>"></td>
<td align=center width=2%><font face="arial" size=2><b><%=rc%>.</td>
<td nowrap><font face="arial" size=2><%=rs.fields("title").value%></td>
<td width=30% align=center><font face="verdana,arial" size=1><%=rs.fields("date").value%></td>
<td width=30% align=right><font face="times new roman" size=2><i><%=rs.fields("author").value%></td>
</tr>
<%
rs.movenext
rc=rc+1
wend
if rc=<1 then
%>
No messages.
<script language="JScript">
function selall()
{
return 0;
}
</script>
<%else%>
<script language="JScript">
function selall()
{
	for(var i=0;i<delform.dmsg.length;i++)
	{
		delform.dmsg[i].checked=delform.chal.checked;
	}
}
</script>
<%
end if
rs.close
con.close
set rs=nothing
set con=nothing
%>
</table>
</form>
</div>
<!--#include virtual="/bbs/codebase.asp"-->
