Attribute VB_Name = "ExportVBAMain"
Option Explicit

' ----------------------------------------------------------------
' Procedure Name: mainExportVBACode
' Purpose: Export ActiveWorkbook VBA code to text files
' Procedure Kind: Sub
' Procedure Access: Public
' Author: gerge
' Date: 15/09/2018
' ----------------------------------------------------------------
Sub mainExportVBACode(control As IRibbonControl)

    Dim srcWb As Workbook
    Dim srcVBAProject As Object
    Dim exportPath As String, exportExtension As String, exportFileName As String
    Dim vbaComponent As Object
    
    Set srcWb = ActiveWorkbook
    Set srcVBAProject = srcWb.vbProject
    
    'check if the VBA project is protected o not
    If f_isVBAProjectProtected(srcVBAProject) = False Then
        
        'export path is equal with the activeworkbook path
        exportPath = srcWb.Path & "\"
        
        'iterate all components in the VBA project
        For Each vbaComponent In srcVBAProject.vbcomponents
            
            'check if there is any code in the component
            If f_isComponentCodeModuleEmpty(vbaComponent) = False Then
                
                'get the text file extension based on component type
                exportExtension = f_getVBAComponentType(vbaComponent)
                
                'if extension is identified export
                If UCase(exportExtension) <> "NONE" Then
                
                    exportFileName = vbaComponent.Name & exportExtension
                    vbaComponent.Export exportPath & exportFileName
                    
                End If
                
            End If
            
        Next vbaComponent
        
    End If

    MsgBox "Export done to folder: " & exportPath
End Sub

