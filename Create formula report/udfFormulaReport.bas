Attribute VB_Name = "udfFormulaReport"
Option Explicit
' ----------------------------------------------------------------
' Procedure Name: f_isThereAnActiveWorkbook
' Purpose: check if there is an ActiveWorkbook to process
' Procedure Kind: Function
' Procedure Access: Public
' Return Type: Boolean
' Author: Gergely Gyetvai
' Date: 23/09/2018
' ----------------------------------------------------------------
Function f_isThereAnActiveWorkbook() As Boolean
    
    Dim userChoice As String

    If Not ActiveWorkbook Is Nothing Then
        f_isThereAnActiveWorkbook = True
    End If

    If f_isThereAnActiveWorkbook = False Then
        MsgBox "There is no ActiveWorkbook to process.", vbCritical, "NO WORKBOOK"
    Else
        userChoice = MsgBox("Formula report is going to be prepared about: " & ActiveWorkbook.Name & vbCrLf & "Would you like to continue?", vbInformation + vbYesNo, "Formula Report")
        If userChoice = vbNo Then
            f_isThereAnActiveWorkbook = False
        End If
    End If
    
End Function

' ----------------------------------------------------------------
' Procedure Name: f_collectFormulaDetails
' Purpose: Collect formula details in a workbook
' Procedure Kind: Function
' Procedure Access: Public
' Parameter srcWb (Workbook): workbook to process
' Return Type: Variant
' Author: Gergely Gyetvai
' Date: 03/11/2018
' ----------------------------------------------------------------
Function f_collectFormulaDetails(srcWb As Workbook) As Variant

    Dim counter As Long
    Dim formulaDetails() As Variant
    Dim cell As Range
    Dim sh As Worksheet

    counter = 0

    For Each sh In srcWb.Worksheets
    
        If f_sheetHasFormula(sh) And f_isSheetProtected(sh) = False Then
        
            'Loop through all cells on the sheet with formula
            For Each cell In sh.Cells.SpecialCells(xlCellTypeFormulas)
            
                'Double check if a cell has formula (required because of merged cells...)
                If cell.HasFormula = True Then
                
                    counter = counter + 1
                    ReDim Preserve formulaDetails(1 To 7, 1 To counter)
                    
                    'Sheet name
                    formulaDetails(1, counter) = sh.Name
                    'Cell address
                    formulaDetails(2, counter) = cell.Address
                    'Cell row
                    formulaDetails(3, counter) = cell.Row
                    'Cell column
                    formulaDetails(4, counter) = cell.Column
                    'Formula value
                    If IsError(cell.Value) Then
                        formulaDetails(5, counter) = "'Error in formula"
                    Else
                        formulaDetails(5, counter) = "'" & cell.Value2
                    End If
                    'Formula
                    formulaDetails(6, counter) = "'" & cell.FormulaLocal
                    'Formula R1C1
                    formulaDetails(7, counter) = "'" & cell.FormulaR1C1Local
                    
                End If
                
            Next cell
        
        End If
    
    Next sh
    
    f_collectFormulaDetails = formulaDetails()

End Function

' ----------------------------------------------------------------
' Procedure Name: f_sheetHasFormula
' Purpose: Examine if a sheet has any formula
' Procedure Kind: Function
' Procedure Access: Public
' Parameter sh (Worksheet): sheet to check
' Return Type: Boolean
' Author: Gergely Gyetvai
' Date: 03/11/2018
' ----------------------------------------------------------------
Function f_sheetHasFormula(sh As Worksheet) As Boolean

    Dim formulaRange As Range

    'Trying to assign formula cells to a range object, if there is no formula cells on
    'the sheet formulaRange remains Nothing
    On Error Resume Next
    Set formulaRange = sh.Cells.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0

    'If formulaRange is Not Nothing there is at least one formula on the sheet
    If Not formulaRange Is Nothing Then
        f_sheetHasFormula = True
    End If
    
    Set formulaRange = Nothing
    
End Function

' ----------------------------------------------------------------
' Procedure Name: f_isSheetProtected
' Purpose: Check if a sheet is protected or not
' Procedure Kind: Function
' Procedure Access: Public
' Parameter sh (Worksheet): sheet to check
' Return Type: Boolean
' Author: Gergely Gyetvai
' Date: 2018. 01. 18.
' ----------------------------------------------------------------
Function f_isSheetProtected(sh As Worksheet) As Boolean

    If sh.ProtectContents = True Then
        f_isSheetProtected = True
    End If

End Function

' ----------------------------------------------------------------
' Procedure Name: f_isArrayEmpty
' Purpose: check if an array is empty
' Procedure Kind: Function
' Procedure Access: Public
' Parameter arrayToCheck (Variant):
' Return Type: Boolean
' Author: http://www.cpearson.com/excel/vbaarrays.htm
' ----------------------------------------------------------------
Function f_isArrayEmpty(arrayToCheck As Variant) As Boolean

    Dim lb As Long
    Dim ub As Long

    Err.Clear
    
    On Error Resume Next
    
    If IsArray(arrayToCheck) = False Then
        ' we weren't passed an array, return True
        f_isArrayEmpty = True
        
    End If

    ' Attempt to get the UBound of the array. If the array is
    ' unallocated, an error will occur.
    ub = UBound(arrayToCheck, 1)
    
    If (Err.Number <> 0) Then
        f_isArrayEmpty = True
    Else
        Err.Clear
        lb = LBound(arrayToCheck)
        If lb > ub Then
            f_isArrayEmpty = True
        Else
            f_isArrayEmpty = False
        End If
    End If

End Function

' ----------------------------------------------------------------
' Procedure Name: f_transposeArray
' Purpose: Transpose a 2D array
' Procedure Kind: Function
' Procedure Access: Public
' Parameter inputArray (Variant): source array
' Return Type: Variant
' Author: Gergely Gyetvai
' Date: 11/01/2019
' ----------------------------------------------------------------
Function f_transposeArray(inputArray As Variant) As Variant

Dim x As Long
Dim y As Long, xUbound As Long
Dim yUbound As Long
Dim tempArray As Variant

    xUbound = UBound(inputArray, 2)
    yUbound = UBound(inputArray, 1)
    
    ReDim tempArray(1 To xUbound, 1 To yUbound)
    
    For x = 1 To xUbound
        For y = 1 To yUbound
            tempArray(x, y) = inputArray(y, x)
        Next y
    Next x
    
    f_transposeArray = tempArray
    
End Function



