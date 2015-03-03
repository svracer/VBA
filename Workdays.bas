Attribute VB_Name = "Workdays"
Option Explicit

Private Function IsWeekend(dtmTemp As Date) As Boolean
  Select Case WeekDay(dtmTemp)
    Case vbSaturday, vbSunday
      IsWeekend = True
    Case Else
      IsWeekend = False
  End Select
   
End Function

Private Function SkipHolidays(rst As Recordset, strField As String, _
        dtmTemp As Date, intIncrement As Integer) As Date
        
  'Skip weekend days, and holidays in the recordset referred to by rst.
  
  Dim strCriteria As String
  On Error GoTo HandleErr
  
  'Move up to the first Monday/last Friday if the first/last
  'of the month was a weekend date. Then skip holidays.
  'Repeat this entire process until you get to a weekday.
  
  Do
    Do While IsWeekend(dtmTemp)
      dtmTemp = dtmTemp + intIncrement
    Loop
    
    If Not rst Is Nothing Then
      If Len(strField) > 0 Then
        If Left(strField, 1) <> "[" Then
          strField = "[" & strField & "]"
        End If
        Do
          strCriteria = strField & "= #" & Format(dtmTemp, "mm/dd/yy") & "#"
          rst.FindFirst strCriteria
          If Not rst.NoMatch Then
            dtmTemp = dtmTemp + intIncrement
          End If
        
        Loop Until rst.NoMatch
      End If
    End If
  Loop Until Not IsWeekend(dtmTemp)
  
ExitHere:
  SkipHolidays = dtmTemp
  Exit Function
  
HandleErr:
  Resume ExitHere
  
End Function

Function NextWorkDay(Optional dtmDate As Date = 0, _
         Optional rst As Recordset = Nothing, _
         Optional strField As String = "") As Date
         
'Return the next working day after the specified date

Dim dtmTemp As Date
Dim strCriteria As String

  If dtmDate = 0 Then
    dtmDate = Date
  End If
  
  NextWorkDay = SkipHolidays(rst, strField, dtmDate + 1, 1)
  
End Function
Function PreviousWorkDay(Optional dtmDate As Date = 0, _
         Optional rst As Recordset = Nothing, _
         Optional strField As String = "") As Date
         
Dim dtmTemp As Date
Dim strCriteria As String

  If dtmDate = 0 Then
    dtmDate = Date
  End If
  
  PreviousWorkDay = SkipHolidays(rst, strField, dtmDate - 1, -1)

End Function

Function FirstWorkdayInMonth(Optional dtmDate As Date = 0, _
        Optional rst As Recordset = Nothing, _
        Optional strField As String = "") As Date

Dim dtmTemp As Date
Dim strCriteria As String

  If dtmDate = 0 Then
    dtmDate = Date
  End If
  
  dtmTemp = DateSerial(Year(dtmDate), Month(dtmDate), 1)
  FirstWorkdayInMonth = SkipHolidays(rst, strField, dtmTemp, 1)
  
End Function

Function LastWorkdayInMonth(Optional dtmDate As Date = 0, _
        Optional rst As Recordset = Nothing, _
        Optional strField As String = "") As Date

Dim dtmTemp As Date
Dim strCriteria As String

  If dtmDate = 0 Then
    dtmDate = Date
  End If
  
  dtmTemp = DateSerial(Year(dtmDate), Month(dtmDate) + 1, 0)
  LastWorkdayInMonth = SkipHolidays(rst, strField, dtmTemp, -1)
  
End Function

Sub TestSkipHolidays()
  Dim rst As DAO.Recordset
  Dim db As DAO.Database
  Set db = DAO.DBEngine.OpenDatabase("C:\My Documents\Holidays.mdb")
  Set rst = db.OpenRecordset("tblHolidays", DAO.dbOpenDynaset)
  
  Debug.Print FirstWorkdayInMonth(#1/19/99#, rst, "date")
  Debug.Print LastWorkdayInMonth(#12/1/99#, rst, "date")
  Debug.Print NextWorkDay(#5/28/99#, rst, "date")
  Debug.Print NextWorkDay(#12/4/98#, rst, "date")
  Debug.Print PreviousWorkDay(#11/27/98#, rst, "date")
  Debug.Print PreviousWorkDay(#12/25/98#, rst, "date")
End Sub




