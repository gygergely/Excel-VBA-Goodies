Attribute VB_Name = "MAINVBAReport"
Option Explicit

Public userCalculationSettings As Variant

' ----------------------------------------------------------------
' Procedure Name: mainVBAProjectReport
' Purpose: Create a report about VBA Project modules & procedures
' Procedure Kind: Sub
' Procedure Access: Public
' Author: Gergely Gyetvai
' Date: 23/09/2018
' ----------------------------------------------------------------
Sub mainVBAProjectReport(control As IRibbonControl)

    Dim srcWb As Workbook
    Dim srcVBAProject As Object
    Dim VBAProjectDetails() As Variant
    
    On Error GoTo errorHandler:
    
    userCalculationSettings = Application.Calculation
    
    Call turnOffThings

    If f_isThereAnActiveWorkbook Then

        Set srcWb = ActiveWorkbook
        Set srcVBAProject = srcWb.VBProject
    
        If f_isVBAProjectProtected(srcVBAProject) = False Then
        
            VBAProjectDetails = f_collectVBAProjectDetails(srcVBAProject)
            
            If f_isArrayEmpty(VBAProjectDetails) Then
                MsgBox "There is no VBA details to report.", vbInformation, "Report"
            Else
                Call createVBAReport(VBAProjectDetails)
            End If
            
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
' Procedure Name: createVBAReport
' Purpose: Creating a report in a new workbook with formatting
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter srcReport (Variant()):
' Author: Gergely Gyetvai
' Date: 23/09/2018
' ----------------------------------------------------------------
Sub createVBAReport(srcReport() As Variant)

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
        .Offset(0, 0) = "Component type"
        .Offset(0, 1) = "VBA Component"
        .Offset(0, 2) = "Procedure Name"
        .Offset(0, 3) = "Procedure Type"
        .Offset(0, 4) = "Total lines"
        .Offset(0, 5) = "Procedure Declaration"
    End With
    
    'Print out array
    sh.Cells(6, 1).Resize(UBound(srcReport, 2), UBound(srcReport, 1)) = Application.WorksheetFunction.Transpose(srcReport)
    
    'Autofit Columns
    For i = 1 To 6
        sh.Columns(i).AutoFit
    Next i
    
    'Add report header to the new workbook
    sh.Cells(1, 1) = "VBA report for " & reportHeader

    'Add generate details
    sh.Cells(2, 1) = "Generated: " & Now() & " by " & Environ("UserName")
    
    'Format columns headers
    Set rng = sh.Range(Cells(5, 1).Address, Cells(5, 6).Address)
    
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
