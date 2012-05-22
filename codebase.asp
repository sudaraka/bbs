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

sub out3d(msg,x,y)
  for i=0 to 2
    response.write("<div class=3d_text"&i&" style=""top:"&y+4-(i*2)&";left:"&x+4-(i*2)&";"">"&msg&"</div>")
  next
end sub

sub loginbox(actfil,size)
  response.write("<form method=""POST"" action="""&actfil&"""><table bgcolor=#6699cc bordercolordark=#000000 bordercolorlight=#99ccff class=blue_table><tr><td align=right><font face=""arial,verdana"" size=2><B>Username:</td><td><input type=""text"" size="""&size&""" name=""username"" class=blue_textbox></td><tr><td align=right><font face=""arial,verdana"" size=2><B>Password:</td><td><input type=""password"" size="""&size&""" name=""password"" class=blue_textbox></td></tr><tr><td></td><td><input type=""submit"" value=""Login"" class=blue_button style=""width:83;""></td></tr></table></form>")
end sub

function splitstr(basestr,ch)
	tmp=instr(basestr,ch)
	if tmp=0 then
		splitstr=basestr
		basestr=""
	else
		splitstr=left(basestr,tmp-1)
		basestr=Mid(basestr, tmp + Len(ch), Len(basestr))
	end if
end function
%>
