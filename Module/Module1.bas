Attribute VB_Name = "Module1"
'Convert Jalali date to Gregorian date.
'If this code works, it is written by Little rabbit.
'If not, it is probably my fault to play with it.

'Vahid Sarfi May 2013

Option Explicit

Private Const mcDayOff = 226894

Private mvarGDayTab
Private mvarJDayTab
Private mcSolar As Double

Public Sub GetJalaliDate(ByVal vGYear As Integer, ByVal vGMonth As Integer, ByVal vGDay As Integer, pJYear As Integer, pJMonth As Integer, pJDay As Integer, pDayName As String)

        
    Dim mGTotalDay As Long
   
    SetConstants
    
    mGTotalDay = GetDayFromFirstGregorianDay(vGYear, vGMonth, vGDay)
    pDayName = GetWeekDayName(mGTotalDay)
    GetJalaliYearMonthDay mGTotalDay, vGYear, vGMonth, vGDay
    pJDay = vGDay
    pJMonth = vGMonth
    pJYear = vGYear
End Sub

Private Sub SetConstants()
    
    mvarGDayTab = Array(Array(0, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31), Array(0, 31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31))
    mvarJDayTab = Array(Array(0, 31, 31, 31, 31, 31, 31, 30, 30, 30, 30, 30, 29), Array(0, 31, 31, 31, 31, 31, 31, 30, 30, 30, 30, 30, 30))
    mcSolar = 365.25 - 0.25 / 33
    
End Sub

Private Function GetDayFromFirstGregorianDay(ByVal vGYaer As Integer, ByVal vGMonth As Integer, ByVal vGDay As Integer) As Long
    
    Dim mGYearDiv4 As Integer, mGYearDiv100 As Integer, mGYearDiv400 As Integer
    Dim mGTotalDays As Long
    
    mGYearDiv4 = vGYaer \ 4
    mGYearDiv100 = vGYaer \ 100
    mGYearDiv400 = vGYaer \ 400
    
    mGTotalDays = GetGDayFromBeginOfYear(vGYaer, vGMonth, vGDay)
    mGTotalDays = CLng(vGYaer - 1) * 365 + mGTotalDays + mGYearDiv4 - mGYearDiv100 + mGYearDiv400
    
    GetDayFromFirstGregorianDay = mGTotalDays
End Function

Private Function GetGDayFromBeginOfYear(ByVal vGYear As Integer, ByVal vGMonth As Integer, ByVal vGDay As Integer) As Long
    Dim mGLeap As Integer
    Dim mCount As Integer
    
    GetGDayFromBeginOfYear = vGDay
    mGLeap = IsLeapGregorian(vGYear)
    For mCount = 1 To vGMonth - 1
        GetGDayFromBeginOfYear = GetGDayFromBeginOfYear + mvarGDayTab(mGLeap)(mCount)
    Next mCount
    
End Function

Private Function IsLeapGregorian(ByVal vGYear As Integer) As Integer

    If (vGYear Mod 4 = 0 And vGYear Mod 100 <> 0) Or (vGYear Mod 400 = 0) Then
        IsLeapGregorian = 1
    Else
        IsLeapGregorian = 0
    End If
End Function

Private Function GetJalaliYearMonthDay(vGTotalDay As Long, pJYear As Integer, pJMonth As Integer, pJDay As Integer)
    
    Dim mJTotalDay As Long
    Dim mJYear As Integer
    Dim mJDay As Integer
    Dim mJLeaps As Integer
    
    mJTotalDay = vGTotalDay - mcDayOff
    mJYear = mJTotalDay \ mcSolar
    
    mJLeaps = GetAllJalaliLeapFromBegin(mJYear)
    
    mJDay = mJTotalDay - (365 * CLng(mJYear) + mJLeaps)
    mJYear = mJYear + 1

    Do While mJDay <= 0
        mJYear = mJYear - 1
        If IsLeapJalali(mJYear) = 1 Then
            mJDay = mJDay + 366
        Else
            mJDay = mJDay + 365
        End If
    Loop
        
    If (mJDay = 366 And IsLeapJalali(mJYear) = 0) Then
        mJDay = 1
        mJYear = mJYear + 1
    End If
    pJYear = mJYear
    GetJalaliMonthDay mJYear, mJDay, pJMonth, pJDay
    
End Function

Private Function IsLeapJalali(ByVal vJYear As Integer) As Integer
    
    Dim mTemp As Integer
    
    mTemp = vJYear Mod 33
    If mTemp = 1 Or mTemp = 5 Or mTemp = 9 Or mTemp = 13 Or mTemp = 17 Or mTemp = 22 Or mTemp = 26 Or mTemp = 30 Then
        IsLeapJalali = 1
    Else
        IsLeapJalali = 0
    End If
End Function

Private Function GetAllJalaliLeapFromBegin(ByVal vJYear As Integer) As Integer

    Dim mJLeap As Integer
    Dim mCurrentCycle As Integer
    Dim mJDiv33 As Integer
    Dim mCount As Integer
    Dim mTemp As Integer
    
    mJDiv33 = vJYear \ 33
    mCurrentCycle = vJYear - (mJDiv33 * 33)
    mJLeap = mJDiv33 * 8
    If mCurrentCycle > 0 Then
        mTemp = IIf(mCurrentCycle <= 18, mCurrentCycle, 18)
        For mCount = 1 To mTemp Step 4
            mJLeap = mJLeap + 1
        Next
    End If
    
    If mCurrentCycle > 21 Then
        mTemp = IIf(mCurrentCycle <= 30, mCurrentCycle, 30)
        For mCount = 22 To mTemp Step 4
            mJLeap = mJLeap + 1
        Next
    End If
    GetAllJalaliLeapFromBegin = mJLeap

End Function


Private Sub GetJalaliMonthDay(ByVal vJYear As Integer, ByVal vJDayOfYear As Integer, pJMonth As Integer, pJDay As Integer)
    Dim mCount As Integer
    Dim mJLeap As Integer

    mJLeap = IsLeapJalali(vJYear)
    mCount = 1
    Do While vJDayOfYear > mvarJDayTab(mJLeap)(mCount)
        vJDayOfYear = vJDayOfYear - mvarJDayTab(mJLeap)(mCount)
        mCount = mCount + 1
    Loop
    pJMonth = mCount
    pJDay = vJDayOfYear
End Sub



Private Function GetWeekDayName(DayFromBegin As Long) As String
    Dim Temp As Integer
    
    Temp = DayFromBegin Mod 7
    Select Case Temp
    
    Case 0
        GetWeekDayName = "Ìﬂ ‘‰»Â"
    Case 1
        GetWeekDayName = "œÊ ‘‰»Â"
    Case 2
        GetWeekDayName = "”Â ‘‰»Â"
    Case 3
        GetWeekDayName = "çÂ«— ‘‰»Â"
    Case 4
        GetWeekDayName = "Å‰Ã ‘‰»Â"
    Case 5
        GetWeekDayName = "Ã„⁄Â"
    Case 6
        GetWeekDayName = "‘‰»Â"
    End Select
    
End Function

Public Sub GetGregorianDate(ByVal vJYear As Integer, ByVal vJMonth As Integer, ByVal vJDay As Integer, ByRef pGYear As Integer, ByRef pGMonth As Integer, ByRef pGDay As Integer, pDayName As String)
    
    Dim mJTotalDays As Long
    Dim mGYear As Integer
    Dim mGMonth As Integer
    Dim mGDay As Integer
    
    SetConstants
    
    mJTotalDays = GetDayFromFirstJalaliDay(vJYear, vJMonth, vJDay)
    GetWeekDayName (mJTotalDays + mcDayOff)
    GetGregorianYearMonthDay mJTotalDays, mGYear, mGMonth, mGDay
    pGYear = mGYear
    pGMonth = mGMonth
    pGDay = mGDay
End Sub

Private Function GetDayFromFirstJalaliDay(ByVal vJYear As Integer, ByVal vJMonth As Integer, ByVal vJDay As Integer) As Long

    Dim mJLeap As Integer
    Dim mTemp As Integer

    mJLeap = GetAllJalaliLeapFromBegin(vJYear - 1)
    mTemp = GetJDayFromBeginOfYear(vJYear, vJMonth, vJDay)
    GetDayFromFirstJalaliDay = CLng((vJYear - 1)) * 365 + mJLeap + mTemp

End Function

Private Function GetJDayFromBeginOfYear(ByVal vJYear As Integer, ByVal vJMonth As Integer, ByVal vJDay As Integer) As Integer

    Dim mCount As Integer
    Dim mJLeap As Integer
    
    GetJDayFromBeginOfYear = vJDay
    mJLeap = IsLeapJalali(vJYear)
    For mCount = 1 To vJMonth - 1
        GetJDayFromBeginOfYear = GetJDayFromBeginOfYear + mvarJDayTab(mJLeap)(mCount)
    Next mCount

End Function

Private Sub GetGregorianYearMonthDay(vJTotalDays As Long, pGYear As Integer, pGMonth As Integer, pGDay As Integer)
    
    Dim mGTotalDays As Long

    Dim mGDiv4 As Integer
    Dim mGDiv100 As Integer
    Dim mGDiv400 As Integer
    Dim mGDays As Integer
    
    mGTotalDays = vJTotalDays + mcDayOff
    pGYear = mGTotalDays \ mcSolar
    mGDiv4 = pGYear \ 4
    mGDiv100 = pGYear \ 100
    mGDiv400 = pGYear \ 400
    
    ' Find Gregorian day of year
    mGDays = mGTotalDays - (365 * CLng(pGYear)) - (mGDiv4 - mGDiv100 + mGDiv400)
    pGYear = pGYear + 1
    
    Do While mGDays <= 0
        pGYear = pGYear - 1
        If IsLeapGregorian(pGYear) = 1 Then
            mGDays = mGDays + 366
        Else
            mGDays = mGDays + 365
        End If
    Loop
    
    If (mGDays = 366 And IsLeapGregorian(pGYear) = 0) Then
        mGDays = 1
        pGYear = pGYear + 1
    End If
    GetGregorianMonthDay pGYear, mGDays, pGMonth, pGDay
End Sub

Private Sub GetGregorianMonthDay(ByVal vGYear As Integer, ByVal vGDayOfYear As Integer, pGMonth As Integer, pGDay As Integer)
    Dim mCount As Integer
    Dim mGLeap
    
    mGLeap = IsLeapGregorian(vGYear)
    mCount = 1
    Do While vGDayOfYear > mvarGDayTab(mGLeap)(mCount)
        vGDayOfYear = vGDayOfYear - mvarGDayTab(mGLeap)(mCount)
        mCount = mCount + 1
    Loop
    pGMonth = mCount
    pGDay = vGDayOfYear
End Sub


