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
if request.form("title")="" and request.form("message")="" then
%>
<html>
<head>
<title> BBS - Sign Up failure</title>
<link href="/<%=session("homepath")%>main.css" rel=stylesheet>
</head>
<body vlink=#0000ff><font face="arial,verdana" size=2>
<font face="times new roman" size=7><i><%out3d "Bulletin Board Service.",0,0%></i></font>
&nbsp;<p><br>&nbsp;&nbsp;&nbsp;
<font color=#cc0000><b>Both <i>title</i> and <i>message</i> fields are empty.</b></font>
<p><a href="/<%=session("homepath")%>compose.asp"><b>Go Back</b></a>
</body>
</html>
<%
	response.end
end if
vtitle=""
tmpstr=request.form("title")
while instr(tmpstr,"'")>0
	vtitle=vtitle&splitstr(tmpstr,"'")&"''"
wend
vtitle=vtitle&tmpstr
if vtitle="" then
	vtitle="-- No title --"
end if
vauthor=request.form("author")
if vauthor="" then
	vauthor="Unknown"
end if
vownerlst=""
tmpstr=lcase(request.form("owner"))
while tmpstr<>""
  vownerlst=vownerlst&splitstr(tmpstr," ")
wend
vmsg=""
tmpstr=request.form("message")
if request.form("sign") then
	tmpstr=tmpstr&"<br>"&session("sign")
end if
while instr(tmpstr,"'")>0
	vmsg=vmsg&splitstr(tmpstr,"'")&"''"
wend
vmsg=vmsg&tmpstr
set con=server.createobject("adodb.connection")
con.open "driver={microsoft access driver (*.mdb)};dbq="&server.mappath("/"&session("dbpath")&"bulletin.mdb")
if lcase(vownerlst)="all@bbs" then
	set rs=con.execute("select username from profiles")
	vownerlst=""
	while not rs.eof
		vownerlst=vownerlst&","&rs.fields("username").value
		rs.movenext
	wend
	rs.close
	set rs=nothing
end if
do
	vowner=splitstr(vownerlst,",")
	doex=true
	if vowner="" then
		vowner="public_bbs"
	end if
	if vowner<>"public_bbs" then
		set rs=con.execute("select * from profiles where username='"&vowner&"'")
		if rs.eof then
			doex=false
		end if
		rs.close
	end if
	if doex then
		con.execute("insert into bbsdata (owner,author,author_ip,title,message) values('"&vowner&"', '"&vauthor&"', '"&request.servervariables("REMOTE_ADDR")&"', '"&vtitle&"', '"&vmsg&"')")
	end if
loop until vownerlst=""
con.close
set rs=nothing
set con=nothing
response.redirect "/"&session("homepath")
%>
<!--#include virtual="/bbs/codebase.asp"-->
