<div align="center">

## Get WeekNumber for a given date


</div>

### Description

This snippet will allow you to pass a string as a date, and get the week number for that calendar year, returned as an integer
 
### More Info
 
Date, passed as a string

Week number in calendar year, returned as an integer


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jeff Mayes](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jeff-mayes.md)
**Level**          |Beginner
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Math/ Dates](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/math-dates__1-37.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jeff-mayes-get-weeknumber-for-a-given-date__1-67573/archive/master.zip)





### Source Code

```
Public Function WeekNum(sDate As String) As Integer
Dim iYear As Integer
Dim iMon As Integer
Dim sTemp As String
Dim iDay11 As Integer
Dim iDiff As Integer
  If IsDate(sDate) And sDate <> "12:00:00 PM" Then
      'determine the year
    iYear = Year(sDate)
      'determine the month
    iMon = Month(sDate)
      'determine weekday of Jan 1st
    sTemp = "1/1/" & Trim(Str$(iYear))
    iDay11 = Weekday(sTemp)
      'now calculate the difference in days
    iDiff = DateDiff("d", sTemp, sDate)
      'base week
    WeekNum = Int(iDiff / 7) + 1
      'check for rollover based on day of week
    If Weekday(sDate) < iDay11 Then
      WeekNum = WeekNum + 1
    End If
  End If
End Function
```

