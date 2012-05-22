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

if session("user")="" or session("acclvl")=0 or session("user")="public_bbs" then
	response.write("<font face=""verdana"" size=1><b>ERROR:</b> Session profile is not available.<br>Please contact <a href=""mailto:sys@post.com?subject=BBS Error""> Systems</a>.<hr>")
	response.end
end if
vdmlst=""
tmpstr=request.form("dmsg")
while tmpstr<>""
  vdmlst=vdmlst&splitstr(tmpstr," ")
wend
if vdmlst<>"" then
	set con=server.createobject("adodb.connection")
	con.open "driver={microsoft access driver (*.mdb)};dbq="&server.mappath("/"&session("dbpath")&"bulletin.mdb")
	while vdmlst<>""
		vdm=splitstr(vdmlst,",")
		con.execute("delete * from bbsdata where id="&vdm&"")
	wend
	con.close
	set con=nothing
end if
response.redirect "/"&session("homepath")
%>
<!--#include virtual="/bbs/codebase.asp"-->
