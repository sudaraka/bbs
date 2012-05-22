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
select case request.form("formtype")
	case "rights"
		if(session("acclvl")and 4)=4 then
			cur=0
			vcurlst=""
			tmpstr=request.form("acr")
			while tmpstr<>""
			  vcurlst=vcurlst&splitstr(tmpstr," ")
			wend
			while vcurlst<>""
				cur=cur+splitstr(vcurlst,",")
			wend
			set con=server.createobject("adodb.connection")
			con.open "driver={microsoft access driver (*.mdb)};dbq="&server.mappath("/"&session("dbpath")&"bulletin.mdb")
			con.execute("update profiles set acclvl="&cur&" where username='"&request.form("vuser")&"'")
			con.close
			set con=nothing
			if lcase(session("user"))=lcase(request.form("vuser"))then session("acclvl")=cur
		end if
	case "profile"
		set con=server.createobject("adodb.connection")
		con.open "driver={microsoft access driver (*.mdb)};dbq="&server.mappath("/"&session("dbpath")&"bulletin.mdb")
		if(session("acclvl")and 2)=2 or (session("acclvl")and 64)=64 then
			con.execute("update profiles set password='"&request.form("pw")&"', fullname='"&request.form("fname")&"', address='"&request.form("address")&"', city='"&request.form("city")&"', zip='"&request.form("zip")&"', country='"&request.form("country")&"', phone='"&request.form("phone")&"', email='"&request.form("email")&"', sign='"&request.form("sign")&"' where username='"&request.form("vuser")&"'")
			if lcase(session("user"))=lcase(request.form("vuser"))then
				session("pw")=request.form("pw")
			end if
		else
			con.execute("update profiles set fullname='"&request.form("fname")&"', address='"&request.form("address")&"', city='"&request.form("city")&"', zip='"&request.form("zip")&"', country='"&request.form("country")&"', phone='"&request.form("phone")&"', email='"&request.form("email")&"', sign='"&request.form("sign")&"' where username='"&request.form("vuser")&"'")
		end if
		con.close
		set con=nothing
		if lcase(session("user"))=lcase(request.form("vuser"))then
			session("fullname")=request.form("fname")
			session("sign")=request.form("sign")
		end if
end select
response.redirect "config.asp"
%>
<!--#include virtual="/bbs/codebase.asp"-->
