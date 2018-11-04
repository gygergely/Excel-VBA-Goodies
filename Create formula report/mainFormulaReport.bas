Attribute VB_Name = "mainFormulaReport"
Option Explicit

Public userCalculationSettings As Variant

' ----------------------------------------------------------------
' Procedure Name: mainFormulaReport
' Purpose: Create report about ActiveWorkbook formulas
' Procedure Kind: Sub
' Procedure Access: Public
' Author: Gergely Gyetvai
' Date: 03/11/2018
' ----------------------------------------------------------------
Sub mainFormulaReportCompile()

    Dim srcWb As Workbook
    Dim formulaDetails() As Variant
    
    On Error GoTo errorHandler:
    
    userCalculationSettings = Application.Calculation
    
    Call turnOffThings
    
    If f_isThereAnActiveWorkbook Then

        Set srcWb = ActiveWorkbook
        
        Call getProtectedSheetNames(srcWb)
        
        formulaDetails = f_collectFormulaDetails(srcWb)
            
        If f_isArrayEmpty(formulaDetails) Then
            MsgBox "There is no formula details to report.", vbInformation, "Report"
        Else
            Call createFormulaReport(formulaDetails)
        End If

    End If
    
exitProcedure:
    Call turnOnThings
    Exit Sub
    
errorHandler:
    MsgBox "Error occured during report creation. " & vbCrLf & Err.Number & " - " & Err.Description, vbCritical, "ERROR"
    On Error GoTo 0
    Resume exitProcedure

End Sub

' ----------------------------------------------------------------
' Procedure Name: createFormulaReport
' Purpose: Creating a report in a new workbook with formatting
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter srcReport (Variant()):
' Author: Gergely Gyetvai
' Date: 03/11/2018
' ----------------------------------------------------------------
Sub createFormulaReport(srcReport() As Variant)

    Dim reportHeader As String
    Dim reportWb As Workbook
    Dim sh As Worksheet
    Dim i As Long
    Dim rng As Range

    'CREATE REPORT
    '================================================

    'Report header full name of the workbook
    reportHeader = ActiveWorkbook.FullName
    
    'Add new workbook
    Set reportWb = Workbooks.Add
    Set sh = reportWb.Worksheets(1)
    
    'Add column headers
    With sh.Cells(5, 1)
        .Offset(0, 0) = "Sheet name"
        .Offset(0, 1) = "Address"
        .Offset(0, 2) = "Row"
        .Offset(0, 3) = "Column"
        .Offset(0, 4) = "Value"
        .Offset(0, 5) = "Formula"
        .Offset(0, 6) = "Formula R1C1"
    End With
    
    'Print out array
    Set rng = sh.Cells(6, 1).Resize(UBound(srcReport, 2), UBound(srcReport, 1))
    rng = Application.WorksheetFunction.Transpose(srcReport)
    
    'Autofit Columns / max width 80
    For i = 1 To 7
        sh.Columns(i).AutoFit
        If sh.Columns(i).ColumnWidth > 80 Then sh.Columns(i).ColumnWidth = 80
    Next i
    
    'Autofit Rows
    rng.Rows.AutoFit
    
    'Add report header to the new workbook
    sh.Cells(1, 1) = "Formula report for " & reportHeader

    'Add generate details
    sh.Cells(2, 1) = "Generated: " & Now() & " by " & Environ("UserName")
    
    'Format columns headers
    Set rng = sh.Range(Cells(5, 1).Address, Cells(5, 7).Address)
    
    With rng
        .Interior.Color = RGB(84, 130, 53)
        .Font.Color = RGB(255, 255, 255)
        .Font.Size = 9
        .Font.Bold = True
        .RowHeight = rng.RowHeight * 2.5
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    'Freeze pane
    sh.Cells(6, 1).Select
    ActiveWindow.FreezePanes = True
    
    'Format header
    Set rng = sh.Range(Cells(1, 1).Address, Cells(2, 1).Address)
    
    With rng
        .Font.Size = 11
        .Font.Color = RGB(55, 86, 35)
        .Font.Bold = True
    End With

End Sub

' ----------------------------------------------------------------
' Procedure Name: getProtectedSheetNames
' Purpose: Check if there is any protected sheet in a workbook and notify user
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter srcWb (Workbook): workbook to check
' Author: Gergely Gyetvai
' Date: 03/11/2018
' ----------------------------------------------------------------
Sub getProtectedSheetNames(srcWb As Workbook)

    Dim nrOfProtectedSheet As Long
    Dim sh As Worksheet
    Dim protectedSheetNames As String

    nrOfProtectedSheet = 0
    
    For Each sh In srcWb.Worksheets
    
        If f_isSheetProtected(sh) = True Then
            nrOfProtectedSheet = nrOfProtectedSheet + 1
            protectedSheetNames = protectedSheetNames & vbCrLf & vbTab & " - " & sh.Name
        End If
        
    Next sh
        
    'If all sheets are protected message to end user
    If nrOfProtectedSheet = srcWb.Sheets.Count Then
    
        MsgBox "All sheets are protected in the workbook." & vbCrLf & "report only works on not protected sheets." & vbCrLf & "Please unprotect the sheet(s).", vbCritical, "All sheets protected."
        
    Else
        'If there is any protected sheet in the workbook, let user know they are out of scope
        'Only check at first load, not at refresh
        If Len(protectedSheetNames) > 0 Then
            MsgBox "The following sheets are protected, therefore they are out of scope: " & protectedSheetNames, vbInformation, "Sheet List out of scope"
        End If
        
    End If
    
End Sub

Sub turnOffThings()

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

End Sub

Sub turnOnThings()

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = userCalculationSettings

End Sub

