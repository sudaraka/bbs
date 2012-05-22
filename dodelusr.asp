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

if session("user")="" or session("acclvl")=0 then
	response.write("<font face=""verdana"" size=1><b>ERROR:</b> Session profile is not available.<br>Please contact <a href=""mailto:sys@post.com?subject=BBS Error""> Systems</a>.<hr>")
	response.end
end if
if(session("acclvl")and 16)=16 or (session("acclvl")and 64)=64 then
	set con=server.createobject("adodb.connection")
	con.open "driver={microsoft access driver (*.mdb)};dbq="&server.mappath("/"&session("dbpath")&"bulletin.mdb")
	con.execute("delete * from profiles where username='"&request.form("delusr")&"'")
	con.execute("delete * from bbsdata where owner='"&request.form("delusr")&"'")
	con.close
	set con=nothing
	if lcase(request.form("delusr"))=lcase(session("user"))then
		response.redirect "/"&session("homepath")&"logout.asp"
	else
		response.redirect "/"&session("homepath")&"config.asp"
	end if
else
	response.redirect "/"&session("homepath")&"config.asp"
end if
%>
