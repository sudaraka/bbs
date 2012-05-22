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
session("errstr")=""
session("homepath")="bbs/"
session("dbpath")="bbs/db/"
%>
<html>
<head>
<title> BBS</title>
<link href="/<%=session("homepath")%>main.css" rel=stylesheet>
</head>
<body vlink=#0000ff><font face="arial,verdana" size=2>
<font face="times new roman" size=7><i><%out3d "Bulletin Board Service.",0,0%></i></font>
<%
if(session("user")="" or lcase(session("user"))="public_bbs")then
session("user")=lcase(request.form("username"))
session("pw")=request.form("password")
end if
	if (session("user")<>"" and lcase(session("user"))<>"public_bbs") then
		set con=server.createobject("adodb.connection")
		con.open "driver={microsoft access driver (*.mdb)};dbq="&server.mappath("/"&session("dbpath")&"bulletin.mdb")
		set rs=con.execute("select * from profiles where username='"&session("user")&"'")
		if rs.eof then
			session("errstr")="User <b>"&session("user")&"</b> is not registered in the BBS."
			session("user")="public_bbs"
			session("pw")="public_bbs"
			session("acclvl")=0
			session("fullname")=""
			session("sign")=""
		else
			if rs.fields("password").value<>session("pw") then
				session("errstr")="Password given for user <b>"&session("user")&"</b> is incorrect."
				session("user")="public_bbs"
				session("pw")="public_bbs"
				session("acclvl")=0
				session("fullname")=""
				session("sign")=""
			else
				session("fullname")=rs.fields("fullname").value
				session("acclvl")=rs.fields("acclvl").value
				session("sign")=rs.fields("sign").value
			end if
		end if
		rs.close
		con.close
		set rs=nothing
		set con=nothing
	else
		session("user")="public_bbs"
		session("pw")="public_bbs"
	end if
%>
<table width=100% style="position:absolute;top:50;left:0;"><tr>
<td width=29%><%
if session("user")="public_bbs" or session("user")="" then
	loginbox "default.asp",10
else
%>
<form name="delform" method="POST" action="/<%=session("homepath")%>delmsg.asp">
<table style="height:86;" width=100% bgcolor=#6699cc bordercolordark=#000000 bordercolorlight=#99ccff class=blue_table><tr><td align=center valign=bottom>
<input style="width:160;" type="submit" value="Delete" class=blue_button>
<input style="width:160;" type="button" value="Configure" class=blue_button onclick="location.href='/<%=session("homepath")%>config.asp';">
<input style="width:160;" type="button" value="Logout" class=blue_button onclick="location.href='/<%=session("homepath")%>logout.asp';">
</td></tr></table>
<%
end if
%></td>
<td valign=top><font face="arial,verdana" size=1>Copyright (c) 2000 Sudaraka Wijesinghe.<br>Created: 07/10/2000.<p>
<center><font size=4>
<%
if session("user")="public_bbs" or session("user")="" then
  response.write("Public ")
else
  response.write(session("fullname")&"'s ")
end if
%>
message board.</font></center>
<%
if session("errstr")<>"" then
  response.write("<font color=#cc0000><i>ERROR: "&session("errstr"))
  session("errstr")=""
end if
%>
</td>
<td width=90 valign=top><table style="height:86;" bgcolor=#6699cc bordercolordark=#000000 bordercolorlight=#99ccff class=blue_table><tr><td valign=bottom>
<%
if session("user")="public_bbs" or session("user")="" then
%>
<input style="width=80;" type="button" value="Sign Up" class=blue_button onclick="location.href='/<%=session("homepath")%>signup.asp'">
<%end if%>
<input style="width=80;" type="button" value="User list" class=blue_button onclick="location.href='/<%=session("homepath")%>userlist.asp'">
<input style="width=80;" type="button" value="Compose" class=blue_button onclick="location.href='/<%=session("homepath")%>compose.asp'">
</td></tr></table></td>
</tr></table>
<div style="position:absolute;top:139;left:3;">
<table class=blue_table width=100% bgcolor=#c0c0c0 bordercolordark=#808080 bordercolorlight=#ffffff cellspacing=0><tr>
<%if (session("user")<>""and session("user")<>"public_bbs") then%>
<td width=2%><input type="checkbox" name="chal" onclick="selall();"></td>
<%end if%>
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
set rs=con.execute("select * from bbsdata where owner='"&session("user")&"'")
rc=1
while not rs.eof
%>
<tr<%if(rc mod 2)=0 then response.write(" bgcolor=#eeeeee")%>>
<%if (session("user")<>""and session("user")<>"public_bbs") then%>
<td align=center width=2%><input type="checkbox" name="dmsg" value="<%=rs.fields("id").value%>"></td>
<%end if%>
<td align=center width=2%><font face="arial" size=2><b><%=rc%>.</td>
<td nowrap><font face="arial" size=2>
<%if rs.fields("message").value<>"" then%>
<a href="/<%=session("homepath")%>view.asp?msgid=<%=rs.fields("id").value%>" style="text-decoration:none;">
<%end if%>
<%=rs.fields("title").value%>
</a>
</td>
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
</body>
</html>
<!--#include virtual="/bbs/codebase.asp"-->
