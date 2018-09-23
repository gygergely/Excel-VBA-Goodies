Attribute VB_Name = "UDFVBAReport"
Option Explicit
' ----------------------------------------------------------------
' Procedure Name: f_collectVBAProjectDetails
' Purpose: Collect VBA procedures details
' Procedure Kind: Function
' Procedure Access: Public
' Parameter srcVBAProject (Object): VBAproject
' Return Type: Variant
' Author: Gergely Gyetvai
' Date: 23/09/2018
' ----------------------------------------------------------------
Function f_collectVBAProjectDetails(srcVBAProject As Object) As Variant

    Dim cdModule As Object
    Dim vbaComponent As Object
    Dim counter As Long
    Dim codeStartLine As Long
    Dim codeTotalLines As Long
    Dim projectDetails() As Variant
    Dim procedureDecLineNumber As Long

    'Iterate through VBA project components
    counter = 0
    For Each vbaComponent In srcVBAProject.vbComponents
    
        'check if there is any code in the component
        If f_isComponentCodeModuleEmpty(vbaComponent) = False Then
            Set cdModule = vbaComponent.codeModule
            
            codeTotalLines = cdModule.CountOfLines
            
            'Loop through procedures
            With cdModule
                codeStartLine = .CountOfDeclarationLines + 1
                
                Do Until codeStartLine >= .CountOfLines
                    counter = counter + 1
                    ReDim Preserve projectDetails(1 To 6, 1 To counter)
                    
                    'Component type
                    projectDetails(1, counter) = f_getModuleType(vbaComponent)
                    'VB component name
                    projectDetails(2, counter) = vbaComponent.Name
                    'Procedure name
                    projectDetails(3, counter) = .ProcOfLine(codeStartLine, 0)
                    procedureDecLineNumber = .procBodyline(projectDetails(3, counter), 0)
                    'Procedure type (Function or sub)
                    projectDetails(4, counter) = f_getProcedureType(.Lines(procedureDecLineNumber, 1))
                    'Procedure total line
                    projectDetails(5, counter) = .ProcCountLines(CStr(projectDetails(3, counter)), 0)
                    'Procedure declaration line
                    projectDetails(6, counter) = .Lines(procedureDecLineNumber, 1)
                    
                    codeStartLine = codeStartLine + projectDetails(5, counter)
                    
                Loop
            End With
        End If
    Next vbaComponent

    f_collectVBAProjectDetails = projectDetails()

End Function

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
        userChoice = MsgBox("VBA Project report is going to be prepared about: " & ActiveWorkbook.Name & vbCrLf & "Would you like to continue?", vbInformation + vbYesNo, "VBA Report")
        If userChoice = vbNo Then
            f_isThereAnActiveWorkbook = False
        End If
    End If
    
End Function

' ----------------------------------------------------------------
' Procedure Name: f_isVBAProjectProtected
' Purpose: Check if a VBA Project is protected or not
' Procedure Kind: Function
' Procedure Access: Public
' Parameter srcVBAProject (Object): VBA project to check
' Return Type: Boolean
' Author: Gergely Gyetvaily Gyetvai
' Date: 15/09/2018
' ----------------------------------------------------------------
Function f_isVBAProjectProtected(srcVBAProject As Object) As Boolean

    If srcVBAProject.Protection = 1 Then f_isVBAProjectProtected = True

    If f_isVBAProjectProtected Then MsgBox "VBA Project is protected report creation is not possible."
    
End Function

' ----------------------------------------------------------------
' Procedure Name: f_isComponentCodeModuleEmpty
' Purpose: Check if a VBA component's code module holds any code
' Procedure Kind: Function
' Procedure Access: Public
' Parameter vbaComponent (Object): component to check
' Return Type: Boolean
' Author: Gergely Gyetvaily Gyetvai
' Date: 15/09/2018
' ----------------------------------------------------------------
Function f_isComponentCodeModuleEmpty(vbaComponent As Object) As Boolean

    If vbaComponent.codeModule.CountOfLines < 3 Then f_isComponentCodeModuleEmpty = True
        
End Function

' ----------------------------------------------------------------
' Procedure Name: f_getModuleType
' Purpose: Enumerate module type
' Procedure Kind: Function
' Procedure Access: Public
' Parameter vbaComponent (Object): VBA component
' Return Type: String
' Author: Gergely Gyetvai
' Date: 23/09/2018
' ----------------------------------------------------------------
Function f_getModuleType(vbaComponent As Object) As String

    Select Case vbaComponent.Type
    
        Case 1
            f_getModuleType = "Standard Module"
        Case 2
            f_getModuleType = "Class Module"
        Case 3
            f_getModuleType = "Form"
        Case 11
            f_getModuleType = "Designer"
        Case 100
            f_getModuleType = "Document Module"
        Case Else
            f_getModuleType = "Unknown"
            
    End Select
    
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
' Procedure Name: f_getProcedureType
' Purpose: Define type based on procedure declaration line
' Procedure Kind: Function
' Procedure Access: Public
' Parameter declarationLine (String):
' Return Type: String
' Author: Gergely Gyetvai
' Date: 23/09/2018
' ----------------------------------------------------------------
Function f_getProcedureType(declarationLine As String) As String

    If Left(declarationLine, 8) = "Function" Or InStr(1, declarationLine, " Function ", vbBinaryCompare) > 0 Then
        f_getProcedureType = "Function"
    Else
        f_getProcedureType = "Sub"
    End If

End Function
