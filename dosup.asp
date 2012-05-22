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

if session("user")="" then
	response.write("<font face=""verdana"" size=1><b>ERROR:</b> Session profile is not available.<br>Please contact <a href=""mailto:sys@post.com?subject=BBS Error""> Systems</a>.<hr>")
	response.end
end if
vwelmsg="<font face=""times new roman"" color=#004080 size=6><i><u>Welcome to  BBS.</u></i></font><p><font size=2 face=""Verdana"">                We would like to thank you for joining the fast growing commiunity of  BBS. You can use this BBS to commiunicate with other members, exchange information as you like.<br> BBS is accessible from any corner of the globe, at any time.<p>       If you ever wanted any help on using the BBS, you can either compose a message to <b>support</b> from the BBS or send an e-mail to <a href=""mailto:sys@post.com""> Systems</a>.<p>Thank you,<br>     BBS Administrator."
set con=server.createobject("adodb.connection")
con.open "driver={microsoft access driver (*.mdb)};dbq="&server.mappath("/"&session("dbpath")&"bulletin.mdb")
erst=""
vuname=lcase(request.form("uname"))
if vuname="" then
	erst=erst&"User name not specified.<br>"
else
	set rs=con.execute("select * from profiles where username='"&vuname&"'")
	if not rs.eof then
		erst=erst&"User "&vuname&" already exist.<br>"
	end if
	rs.close
	set rs=nothing
end if
if request.form("pw1")="" then
	erst=erst&"Password not specified.<br>"
end if
if  request.form("pw2")="" then
		erst=erst&"Password not confirmed.<br>"
end if
if request.form("pw1")<>request.form("pw2") then
	erst=erst&"Two password given are not the same.<br>"
end if
vfname=request.form("fname")
if vfname="" then
	erst=erst&"Full name not specified.<br>"
end if
if erst="" then
	con.execute("insert into profiles (username,password,fullname,address,city,zip,country,acclvl,phone,email,suip,sudate) values('"&vuname&"','"&request.form("pw1")&"','"&vfname&"','"&request.form("addr")&"','"&request.form("city")&"','"&request.form("zip")&"','"&request.form("country")&"',34,'"&request.form("phone")&"','"&request.form("email")&"','"&request.servervariables("REMOTE_HOST")&"','"&now&"')")
	con.execute("insert into bbsdata (owner,author,author_ip,title,message) values('"&vuname&"','BBS','127.0.0.1','Welcome to  BBS...','"&vwelmsg&"')")
else
%>
<html>
<head>
<title> BBS - Sign Up failure</title>
<link href="/<%=session("homepath")%>main.css" rel=stylesheet>
</head>
<body vlink=#0000ff><font face="arial,verdana" size=2>
<font face="times new roman" size=7><i><%out3d "Bulletin Board Service.",0,0%></i></font>
&nbsp;<p><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
Following error(s) occured while processing you form.<p><font color=#cc0000><b>
<%=erst%></b></font>
<p>Click Back button of the browser and correct those errors.
</body>
</html>
<%
end if
con.close
set con=nothing
if erst="" then response.redirect "/"&session("homepath")
%>
<!--#include virtual="/bbs/codebase.asp"-->
