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
	response.write("<font face=""verdana"" size=1><b>ERROR:</b> Session profile is not available. Or you don't have access to this section of the site.<br>Please contact <a href=""mailto:sys@post.com?subject=BBS Error""> Systems</a>.<hr>")
	response.end
end if
%>
<html>
<head>
<title> BBS - Configuration</title>
<link href="/<%=session("homepath")%>main.css" rel=stylesheet>
</head>
<body vlink=#0000ff><font face="arial,verdana" size=2>
<font face="times new roman" size=7><i><%out3d "Bulletin Board Service.",0,0%></i></font>
<div style="position:absolute;top:55;">
<form method="POST" action="/<%=session("homepath")%>config.asp">
<table width=100% bgcolor=#6699cc bordercolordark=#000000 bordercolorlight=#99ccff class=blue_table>
<tr><td align=center>
<select name="selusr" style="font-family:arial;font-weight:bold;width=150;">
<%
vusername=request.form("selusr")
if vusername="" then
	vusername=lcase(session("user"))
end if
set con=server.createobject("adodb.connection")
con.open "driver={microsoft access driver (*.mdb)};dbq="&server.mappath("/"&session("dbpath")&"bulletin.mdb")
if(session("acclvl")and 8)=8 then
	set rs=con.execute("select * from profiles")
else
	set rs=con.execute("select * from profiles where username='"&session("user")&"'")
end if
rc=1
while not rs.eof
%>
<option value="<%=lcase(rs.fields("username").value)%>"<%
if lcase(rs.fields("username").value)=vusername then
	vfname=rs.fields("fullname").value
	vemail=rs.fields("email").value
	vpw=rs.fields("password").value
	vaddr=rs.fields("address").value
	vcity=rs.fields("city").value
	vzip=rs.fields("zip").value
	vcountry=rs.fields("country").value
	vacclvl=rs.fields("acclvl").value
	vphone=rs.fields("phone").value
	vsign=rs.fields("sign").value
	response.write(" selected")
end if
%>
><%=rc%>. <%=rs.fields("username").value%></option>
<%
rc=rc+1
rs.movenext
wend
rs.close
set rs=nothing
con.close
set con=nothing
%>
</select>
<input type="submit" class=blue_button value="Load profile" style="width:100;"> <input type="button" class=blue_button value="Go Back" onclick="location.href='/<%=session("homepath")%>';">
</td></tr>
</table>
</form>
<form method="POST" action="/<%=session("homepath")%>confupdt.asp"><input type="hidden" name="formtype" value="profile"><input type="hidden" name="vuser" value="<%=vusername%>">
<table width=100% bgcolor=#6699cc bordercolordark=#000000 bordercolorlight=#99ccff class=blue_table>
<th></th><th></th><th align=left><font face="verdana" color=#ffffff size=2><%response.write("Profile of ["&vusername&"]")%></th>
<tr>
<td align=right width=22%><font face="arial,verdana" size=2><b>Full name:</td>
<td width=10></td>
<td><input type="text" size="30" name="fname" class=blue_textbox value="<%=vfname%>"></td>
</tr>
<tr>
<td align=right width=22%><font face="arial,verdana" size=2><b>E-Mail:</td>
<td width=10></td>
<td><input type="text" size="20" name="email" class=blue_textbox value="<%=vemail%>"></td>
</tr>
<%if((session("acclvl")and 2)=2 and (vacclvl and 64)<>64) or (lcase(session("user"))=vusername and (vacclvl and 64)=64)then%>
<tr>
<td align=right width=22%><font face="arial,verdana" size=2><b>Password:</td>
<td width=10></td>
<td><input type="password" size="15" name="pw" class=blue_textbox value="<%=vpw%>"></td>
</tr>
<%end if%>
<tr>
<td align=right width=22%><font face="arial,verdana" size=2><b>Address:</td>
<td width=10></td>
<td><input type="text" size="50" name="address" class=blue_textbox value="<%=vaddr%>"></td>
</tr>
<tr>
<td align=right width=22%><font face="arial,verdana" size=2><b>City:</td>
<td width=10></td>
<td><input type="text" size="15" name="city" class=blue_textbox value="<%=vcity%>"></td>
</tr>
<tr>
<td align=right width=22%><font face="arial,verdana" size=2><b>ZIP/Postal code:</td>
<td width=10></td>
<td><input type="text" size="10" name="zip" class=blue_textbox value="<%=vzip%>"></td>
</tr>
<tr>
<td align=right width=22%><font face="arial,verdana" size=2><b>Country:</td>
<td width=10></td>
<td><input type="text" size="15" name="country" class=blue_textbox value="<%=vcountry%>"></td>
</tr>
<tr>
<td align=right width=22%><font face="arial,verdana" size=2><b>Telephone:</td>
<td width=10></td>
<td><input type="text" size="15" name="phone" class=blue_textbox value="<%=vphone%>"></td>
</tr>
<tr>
<td align=right valign=top width=22%><font face="arial,verdana" size=2><b>Signature:</td>
<td width=10></td>
<td><textarea rows="3" cols="50" name="sign" class=blue_msgbox><%=vsign%></textarea></td>
</tr>
<tr>
<td></td>
<td></td>
<td><input type="submit" class=blue_button value="Update profile" style="width:100;"></td>
</tr>
</table>
</form>
<%if(session("acclvl")and 4)=4 then%>
<form method="POST" action="/<%=session("homepath")%>confupdt.asp"><input type="hidden" name="formtype" value="rights"><input type="hidden" name="vuser" value="<%=vusername%>">
<table width=100% bgcolor=#6699cc bordercolordark=#000000 bordercolorlight=#99ccff class=blue_table>
<th></th><th align=left><font face="verdana" color=#ffffff size=2><%response.write("Access rights of ["&vusername&"]")%></th>
<tr>
<td width=30%><font face="arial" size=2 color=#0><input type="checkbox" name="acr" value="1"<%if(vacclvl and 1)=1 then response.write(" checked")%>>Manage public messages.</td>
<td width=30%><font face="arial" size=2 color=#0><input type="checkbox" name="acr" value="2"<%if(vacclvl and 2)=2 then response.write(" checked")%>>Change password.</td>
<td width=30%><font face="arial" size=2 color=#0><input type="checkbox" name="acr" value="4"<%if(vacclvl and 4)=4 then response.write(" checked")%>>Change access rights.</td>
</tr>
<tr>
<td width=30%><font face="arial" size=2 color=#0><input type="checkbox" name="acr" value="8"<%if(vacclvl and 8)=8 then response.write(" checked")%>>Administor other users.</td>
<td width=30%><font face="arial" size=2 color=#0><input type="checkbox" name="acr" value="16"<%if(vacclvl and 16)=16 then response.write(" checked")%>>Delete accounts.</td>
<td width=30%><font face="arial" size=2 color=#0><input type="checkbox" name="acr" value="32"<%if(vacclvl and 32)=32 then response.write(" checked")%>>Compose private messages.</td>
</tr>
<tr>
<%if(session("acclvl")and 64)=64 and (lcase(session("user"))=vusername or (vacclvl and 64)<>64)then%>
<td width=30%><font face="arial" size=2 color=#0><input type="checkbox" name="acr" value="64"<%if(vacclvl and 64)=64 then response.write(" checked")%>>Hyper access.</td>
<%else%>
<input type="hidden" name="acr" value="64"<%if(vacclvl and 64)=64 then response.write(" checked")%>>
<%end if%>
</tr>
<tr>
<td></td>
<td><input type="submit" class=blue_button value="Update rights" style="width:100;"></td>
<td></td>
</tr>
</table>
</form>
<%end if%>
<table width=100% bgcolor=#6699cc bordercolordark=#000000 bordercolorlight=#99ccff class=blue_table>
<tr><td align=center>
<%if((session("acclvl")and 16)=16 and (vacclvl and 64)<>64)or(lcase(session("user"))=vusername and (vacclvl and 64)=64)then%>
<input type="button" class=blue_button value="Delete user" onclick="location.href='/<%=session("homepath")%>delusr.asp?userid=<%=vusername%>&username=<%=vfname%>';">
<%end if%>
<%if(session("acclvl")and 1)=1 then%>
<input type="button" class=blue_button value="Manage public messages" onclick="location.href='/<%=session("homepath")%>pubbbs.asp';">
<%end if%>
<input type="button" class=blue_button value="Go Back" onclick="location.href='/<%=session("homepath")%>';">
</td></tr></table>
</div>
</body>
</html>
<!--#include virtual="/bbs/codebase.asp"-->
