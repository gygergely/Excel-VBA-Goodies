Attribute VB_Name = "Test"
Option Explicit

' ----------------------------------------------------------------
' Procedure Name: generateRandomNamedRanges
' Purpose: Add 999 random named ranges (A1:Z100)
' Procedure Kind: Sub
' Procedure Access: Public
' Author: Gergely Gyetvai
' Date: 31/12/2018
' ----------------------------------------------------------------
Sub generateRandomNamedRanges()

    Dim i As Long
    Dim nameType As Integer
    Dim rangeStart As String, rangeEnd As String
    Dim nameRange As Range

    Const charLow As Integer = 65
    Const charHigh As Integer = 90
    Const namingConvention As String = "NamedFormula"

    Call turnOffThings
    
    For i = 1 To 999
        nameType = f_randBetween(0, 1)
    
        rangeStart = CStr(Chr(f_randBetween(charLow, charHigh)) & f_randBetween(1, 100))
    
        Select Case nameType
            Case 0
                Set nameRange = NameTest.Range(rangeStart)
            Case 1
                rangeEnd = CStr(Chr(f_randBetween(charLow, charHigh)) & f_randBetween(1, 100))
                Set nameRange = NameTest.Range(rangeStart & " : " & rangeEnd)
        End Select
            
        ThisWorkbook.Names.Add Name:=namingConvention & Right("0000" & i, 4), RefersTo:=nameRange
    Next i
    
    Call turnOnThings
    
End Sub

' ----------------------------------------------------------------
' Procedure Name: clearAllNames
' Purpose: Clear all names from Thisworkbook
' Procedure Kind: Sub
' Procedure Access: Public
' Author: Gergely Gyetvai
' Date: 31/12/2018
' ----------------------------------------------------------------
Sub clearAllNames()

    Dim nm As Name
    
    Call turnOffThings

    For Each nm In ThisWorkbook.Names
        nm.Delete
    Next nm
    
    Call turnOnThings
    
End Sub

' ----------------------------------------------------------------
' Procedure Name: f_randBetween
' Purpose: Returns a random number in a range
' Procedure Kind: Function
' Procedure Access: Public
' Parameter low (Integer): low end of the range
' Parameter high (Integer): high end of the range
' Return Type: Integer
' Author: Gergely Gyetvai
' Date: 30/12/2018
' ----------------------------------------------------------------
Function f_randBetween(low As Integer, high As Integer) As Integer
    
    Randomize
    f_randBetween = Int((high - low + 1) * Rnd + low)
    
End Function

Sub turnOffThings()

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False

End Sub

Sub turnOnThings()

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True

End Sub

