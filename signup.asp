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
%>
<html>
<head>
<title> BBS - Sign Up</title>
<link href="/<%=session("homepath")%>main.css" rel=stylesheet>
</head>
<body vlink=#0000ff><font face="arial,verdana" size=2>
<font face="times new roman" size=7><i><%out3d "Bulletin Board Service.",0,0%></i></font>
&nbsp;<p><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
Fill ut the following form and click the submit button in order to join the  BBS. Be sure you fill all the required fields marked with "<font color=#ff0000>*</font>".<br>
<form action="/<%=session("homepath")%>dosup.asp" method="POST">
<table width=100% bgcolor=#6699cc bordercolordark=#000000 bordercolorlight=#99ccff class=blue_table>
<tr>
<td align=right width=25%><font face="arial,verdana" size=2><B>Username:<font color=#ff0000>*</td>
<td width=10></td>
<td><input type="text" size="15" name="uname" class=blue_textbox></td>
</tr>
<tr>
<td align=right width=25%><font face="arial,verdana" size=2><B>Password:<font color=#ff0000>*</td>
<td width=10></td>
<td><input type="password" size="15" name="pw1" class=blue_textbox></td>
</tr>
<tr>
<td align=right width=25%><font face="arial,verdana" size=2><B>Re-type Password:<font color=#ff0000>*</td>
<td width=10></td>
<td><input type="password" size="15" name="pw2" class=blue_textbox></td>
</tr>
<tr>
<td align=right width=25%><font face="arial,verdana" size=2><B>E-Mail address:</td>
<td width=10></td>
<td><input type="text" size="20" name="email" class=blue_textbox></td>
</tr>
<tr>
<td align=right width=25%><font face="arial,verdana" size=2><B>Full name:<font color=#ff0000>*</td>
<td width=10></td>
<td><input type="text" size="30" name="fname" class=blue_textbox></td>
</tr>
<tr>
<td align=right width=25%><font face="arial,verdana" size=2><B>Address:</td>
<td width=10></td>
<td><input type="text" size="50" name="addr" class=blue_textbox></td>
</tr>
<tr>
<td align=right width=25%><font face="arial,verdana" size=2><B>City:</td>
<td width=10></td>
<td><input type="text" size="15" name="city" class=blue_textbox></td>
</tr>
<tr>
<td align=right width=25%><font face="arial,verdana" size=2><B>ZIP/Postal code:</td>
<td width=10></td>
<td><input type="text" size="10" name="zip" class=blue_textbox></td>
</tr>
<tr>
<td align=right width=25%><font face="arial,verdana" size=2><B>Country:</td>
<td width=10></td>
<td><input type="text" size="15" name="country" class=blue_textbox></td>
</tr>
<tr>
<td align=right width=25%><font face="arial,verdana" size=2><B>Telephone:</td>
<td width=10></td>
<td><input type="text" size="15" name="phone" class=blue_textbox></td>
</tr>
<td align=right width=25%>&nbsp;</td>
<td width=10></td>
<td>&nbsp;</td>
</tr>
<td align=right width=25%></td>
<td width=10></td>
<td><input style="width:100;" type="submit" class=blue_button value="Submit"> <input  style="width:100;" type="reset" class=blue_button value="Reset"> <input type="button" value="Go Back" class=blue_button onclick="location.href='/<%=session("homepath")%>';"></td>
</tr>
</table>
</form>
</body>
</html>
<!--#include virtual="/bbs/codebase.asp"-->
