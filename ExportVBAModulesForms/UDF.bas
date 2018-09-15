Attribute VB_Name = "UDF"
' ----------------------------------------------------------------
' Procedure Name: f_isVBAProjectProtected
' Purpose: Check if a VBA Project is protected or not
' Procedure Kind: Function
' Procedure Access: Public
' Parameter srcVBAProject (Object): VBA project to check
' Return Type: Boolean
' Author: gerge
' Date: 15/09/2018
' ----------------------------------------------------------------
Function f_isVBAProjectProtected(srcVBAProject As Object) As Boolean

    If srcVBAProject.Protection = 1 Then f_isVBAProjectProtected = True

    If f_isVBAProjectProtected Then MsgBox "VBA Project is protected export not possible"
    
End Function

' ----------------------------------------------------------------
' Procedure Name: f_getVBAComponentType
' Purpose: Get file extension based on component type
' Procedure Kind: Function
' Procedure Access: Public
' Parameter vbaComponent (Object): component
' Return Type: String
' Author: Gergely Gyetvai
' Date: 15/09/2018
' ----------------------------------------------------------------
Function f_getVBAComponentType(vbaComponent As Object) As String

    Select Case vbaComponent.Type
        'Class, sheet and Thisworkbook modules
        Case vbext_ct_ClassModule, vbext_ct_Document
            f_getVBAComponentType = ".cls"
            'Userforms
        Case vbext_ct_MSForm
            f_getVBAComponentType = ".frm"
            'Standard module
        Case vbext_ct_StdModule
            f_getVBAComponentType = ".bas"
        Case Else
            f_getVBAComponentType = "None"
    End Select
        
End Function

' ----------------------------------------------------------------
' Procedure Name: f_isComponentCodeModuleEmpty
' Purpose: Check if a VBA component's code module holds any code
' Procedure Kind: Function
' Procedure Access: Public
' Parameter vbaComponent (Object): component to check
' Return Type: Boolean
' Author: Gergely Gyetvai
' Date: 15/09/2018
' ----------------------------------------------------------------
Function f_isComponentCodeModuleEmpty(vbaComponent As Object) As Boolean

    If vbaComponent.CodeModule.CountOfLines < 3 Then f_isComponentCodeModuleEmpty = True
        
End Function

