<%
Dim MyMonth 'Month of calendar
Dim MyYear 'Year of calendar
Dim FirstDay 'First day of the month. 1 = Monday
Dim CurrentDay 'Used to print dates in calendar
Dim Col 'Calendar column
Dim Row 'Calendar row

MyMonth = Request.Querystring("Month")
MyYear = Request.Querystring("Year")

If IsEmpty(MyMonth) then MyMonth = Month(Date)
if IsEmpty(MyYear) then MyYear = Year(Date)

Call ShowHeader (MyMonth, MyYear)

FirstDay = WeekDay(DateSerial(MyYear, MyMonth, 1)) -1
CurrentDay = 1

'Let's build the calendar
For Row = 0 to 5
For Col = 0 to 6
If Row = 0 and Col < FirstDay then
response.write "<td>&nbsp;</td>"
elseif CurrentDay > LastDay(MyMonth, MyYear) then
response.write "<td>&nbsp;</td>"
else
response.write "<td"

if MyMonth = Month(Date) and CurrentDay = Day(Date) then 
response.write " bgcolor='#FF9999' align='center'>"
else 
response.write " align='center'>"
end if

response.write "<font face='Arial, Helvetica, sans-serif' size='1'>" _
& CurrentDay & "</font></td>"

CurrentDay = CurrentDay + 1
End If

Next
response.write "</tr>"
Next

response.write "</table></body></html>"


'------ Sub and functions


Sub ShowHeader(MyMonth,MyYear)
%>
<html>
<head>
<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>
</head>

<body bgcolor='#FFFFFF'>

<table border='1' cellspacing='3' cellpadding='3' width='200' align='center'>
<tr align='center'> 
<td colspan='7'>
<table border='0' cellspacing='1' cellpadding='1' width='100%'>
<tr>
<td align='left'>
<font face="Arial, Helvetica, sans-serif" size="1">

<%
response.write "<a href = 'calendario.asp?"
if MyMonth - 1 = 0 then 
response.write "month=12&year=" & MyYear -1
else 
response.write "month=" & MyMonth - 1 & "&year=" & MyYear
end if
response.write "'><font size='-1'><</font></a>"
%>

</font>

</td><td align='center'>
<font face="Arial, Helvetica, sans-serif" size="1">
<%
response.write "<b>" & MonthName(MyMonth) & " " & MyYear & "</b>"
%>

</font>

</td><td align='right'>
<font face="Arial, Helvetica, sans-serif" size="1">
<%
response.write "<a href = 'calendario.asp?"
if MyMonth + 1 = 13 then 
response.write "month=1&year=" & MyYear + 1
else 
response.write "month=" & MyMonth + 1 & "&year=" & MyYear
end if
response.write "'><font size='-1'>></font></a>"
%>

</font>

</td></tr></table>
</td>
</tr>
<tr align='center'> 
<td><font face='Arial, Helvetica, sans-serif' size="1">
<b><i>D</i></b></font></td>
<td><font face='Arial, Helvetica, sans-serif' size="1">
<b><i>L</i></b></font></td>
<td><font face='Arial, Helvetica, sans-serif' size="1">
<b><i>M</i></b></font></td>
<td><font face="Arial, Helvetica, sans-serif" size="1">
<b><i>X</i></b></font></td>
<td><font face='Arial, Helvetica, sans-serif' size="1">
<b><i>J</i></b></font></td>
<td><font face='Arial, Helvetica, sans-serif' size="1">
<b><i>V</i></b></font></td>
<td><font face='Arial, Helvetica, sans-serif' size="1">
<b><i>S</i></b></font></td>
</tr>
<%
End Sub

Function MonthName(MyMonth)
Select Case MyMonth
Case 1
MonthName = "Enero"
Case 2
MonthName = "Febrero"
Case 3
MonthName = "Marzo"
Case 4
MonthName = "Abril"
Case 5
MonthName = "Mayo"
Case 6
MonthName = "Junio"
Case 7
MonthName = "Julio"
Case 8
MonthName = "Agosto"
Case 9
MonthName = "Septiembre"
Case 10
MonthName = "Octubre"
Case 11
MonthName = "Noviembre"
Case 12
MonthName = "Diciembre"
Case Else
MonthName = "ERROR!"
End Select
End Function

Function LastDay(MyMonth, MyYear)
' Returns the last day of the month. Takes into account leap years
' Usage: LastDay(Month, Year)
' Example: LastDay(12,2000) or LastDay(12) or Lastday


Select Case MyMonth
Case 1, 3, 5, 7, 8, 10, 12
LastDay = 31

Case 4, 6, 9, 11
LastDay = 30

Case 2
If IsDate(MyYear & "-" & MyMonth & "-" & "29") Then LastDay = 29 Else LastDay = 28

Case Else
LastDay = 0

End Select
End Function
%>